"""Unit tests for Outlook Token Manager."""

from __future__ import annotations
import json
import unittest
from unittest.mock import MagicMock, patch
from urllib.error import HTTPError
from outlook_graph.auth.token_manager import OutlookTokenManager
from outlook_graph.types.account import AccountCredentials
from outlook_graph.types.auth import TokenCacheEntry, TokenManagerOptions
from outlook_graph.types.errors import RateLimitError, TokenRefreshError


class TestTokenManager(unittest.TestCase):
    """Test suite for OAuth 2.0 token management, caching, rotation, and retries."""

    def setUp(self) -> None:
        self.mock_account = AccountCredentials(
            email="test@hotmail.com",
            password="password123",
            refresh_token="mock-refresh-token-1",
            client_id="mock-client-id",
        )

    @patch("urllib.request.urlopen")
    def test_successfully_refresh_token_and_update_cache(
        self, mock_urlopen: MagicMock
    ) -> None:
        mock_response_data = {
            "token_type": "Bearer",
            "scope": "https://graph.microsoft.com/Mail.Read offline_access",
            "expires_in": 3600,
            "access_token": "mock-access-token-xyz",
            "refresh_token": "mock-refresh-token-2",
        }

        mock_resp = MagicMock()
        mock_resp.status = 200
        mock_resp.read.return_value = json.dumps(mock_response_data).encode(
            "utf-8"
        )
        mock_resp.__enter__.return_value = mock_resp
        mock_urlopen.return_value = mock_resp

        rotated_events: list[tuple[str, str]] = []

        def on_token_rotated(
            old_rt: str, new_rt: str, acc: AccountCredentials
        ) -> None:
            rotated_events.append((old_rt, new_rt))

        options = TokenManagerOptions(on_token_rotated=on_token_rotated)
        token_manager = OutlookTokenManager(options)

        token = token_manager.get_access_token(self.mock_account)

        self.assertEqual(token, "mock-access-token-xyz")
        self.assertEqual(self.mock_account.refresh_token, "mock-refresh-token-2")
        self.assertEqual(
            rotated_events, [("mock-refresh-token-1", "mock-refresh-token-2")]
        )

        cached = token_manager.get_cached_token(self.mock_account.email)
        self.assertIsNotNone(cached)
        self.assertEqual(cached.access_token, "mock-access-token-xyz")

    @patch("urllib.request.urlopen")
    def test_return_cached_token_if_valid(
        self, mock_urlopen: MagicMock
    ) -> None:
        token_manager = OutlookTokenManager()
        # Set cache with expires_at far in the future
        token_manager._token_cache["test@hotmail.com"] = TokenCacheEntry(
            access_token="valid-cached-token",
            refresh_token="rt",
            expires_at=9999999999999,
            scope="Mail.Read",
        )

        token = token_manager.get_access_token(self.mock_account)
        self.assertEqual(token, "valid-cached-token")
        mock_urlopen.assert_not_called()

    @patch("urllib.request.urlopen")
    def test_force_refresh_ignores_cache(
        self, mock_urlopen: MagicMock
    ) -> None:
        mock_response_data = {
            "token_type": "Bearer",
            "scope": "Mail.Read",
            "expires_in": 3600,
            "access_token": "new-fresh-token",
            "refresh_token": "mock-refresh-token-1",
        }
        mock_resp = MagicMock()
        mock_resp.status = 200
        mock_resp.read.return_value = json.dumps(mock_response_data).encode(
            "utf-8"
        )
        mock_resp.__enter__.return_value = mock_resp
        mock_urlopen.return_value = mock_resp

        token_manager = OutlookTokenManager()
        token_manager._token_cache["test@hotmail.com"] = TokenCacheEntry(
            access_token="old-cached-token",
            refresh_token="rt",
            expires_at=9999999999999,
            scope="Mail.Read",
        )

        token = token_manager.get_access_token(
            self.mock_account, force_refresh=True
        )
        self.assertEqual(token, "new-fresh-token")
        self.assertEqual(mock_urlopen.call_count, 1)

    @patch("urllib.request.urlopen")
    def test_throw_token_refresh_error_on_400_invalid_grant(
        self, mock_urlopen: MagicMock
    ) -> None:
        err_body = json.dumps(
            {
                "error": "invalid_grant",
                "error_description": "AADSTS700082: The refresh token has expired",
            }
        ).encode("utf-8")

        mock_err = HTTPError(
            url="https://login.microsoftonline.com/common/oauth2/v2.0/token",
            code=400,
            msg="Bad Request",
            hdrs=MagicMock(),
            fp=MagicMock(read=MagicMock(return_value=err_body)),
        )
        mock_urlopen.side_effect = mock_err

        token_manager = OutlookTokenManager(
            TokenManagerOptions(max_retries=0, initial_retry_delay_ms=10)
        )

        with self.assertRaises(TokenRefreshError) as ctx:
            token_manager.refresh_access_token(self.mock_account)

        self.assertIn("AADSTS700082", str(ctx.exception))
    @patch("urllib.request.urlopen")
    def test_concurrent_get_access_token_deduplication(
        self, mock_urlopen: MagicMock
    ) -> None:
        mock_response_data = {
            "token_type": "Bearer",
            "scope": "Mail.Read",
            "expires_in": 3600,
            "access_token": "dedup-token-123",
            "refresh_token": "mock-refresh-token-1",
        }
        mock_resp = MagicMock()
        mock_resp.status = 200
        mock_resp.read.return_value = json.dumps(mock_response_data).encode("utf-8")
        mock_resp.__enter__.return_value = mock_resp
        mock_urlopen.return_value = mock_resp

        token_manager = OutlookTokenManager()
        token1 = token_manager.get_access_token(self.mock_account)
        token2 = token_manager.get_access_token(self.mock_account)

        self.assertEqual(token1, "dedup-token-123")
        self.assertEqual(token2, "dedup-token-123")
        self.assertEqual(mock_urlopen.call_count, 1)


if __name__ == "__main__":
    unittest.main()
