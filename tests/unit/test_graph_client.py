"""Unit tests for Graph HTTP Client."""

from __future__ import annotations
import json
import unittest
from unittest.mock import MagicMock, patch
from urllib.error import HTTPError
from outlook_graph.auth.token_manager import OutlookTokenManager
from outlook_graph.http.graph_client import GraphHttpClient, RequestOptions
from outlook_graph.types.account import AccountCredentials
from outlook_graph.types.errors import GraphApiError, RateLimitError


class TestGraphClient(unittest.TestCase):
    """Test suite for resilient GraphHttpClient."""

    def setUp(self) -> None:
        self.mock_account = AccountCredentials(
            email="user@hotmail.com",
            refresh_token="rt",
            client_id="cid",
        )
        self.mock_token_mgr = MagicMock(spec=OutlookTokenManager)
        self.mock_token_mgr.get_access_token.return_value = "mock-bearer-token"
        self.client = GraphHttpClient(
            self.mock_token_mgr, timeout_ms=5000, max_retries=1
        )

    @patch("urllib.request.urlopen")
    def test_successful_json_get(self, mock_urlopen: MagicMock) -> None:
        mock_resp = MagicMock()
        mock_resp.status = 200
        mock_resp.headers = {"Content-Type": "application/json; charset=utf-8"}
        mock_resp.read.return_value = json.dumps({"value": [1, 2, 3]}).encode(
            "utf-8"
        )
        mock_resp.__enter__.return_value = mock_resp
        mock_urlopen.return_value = mock_resp

        res = self.client.get(
            self.mock_account, "https://graph.microsoft.com/v1.0/me/messages"
        )
        self.assertEqual(res, {"value": [1, 2, 3]})

    @patch("urllib.request.urlopen")
    def test_204_no_content_returns_none(
        self, mock_urlopen: MagicMock
    ) -> None:
        mock_resp = MagicMock()
        mock_resp.status = 204
        mock_resp.headers = {}
        mock_resp.read.return_value = b""
        mock_resp.__enter__.return_value = mock_resp
        mock_urlopen.return_value = mock_resp

        res = self.client.delete(
            self.mock_account, "https://graph.microsoft.com/v1.0/me/messages/123"
        )
        self.assertIsNone(res)

    @patch("urllib.request.urlopen")
    def test_401_triggers_token_refresh_and_retry(
        self, mock_urlopen: MagicMock
    ) -> None:
        # First call fails with 401
        err_401 = HTTPError(
            url="https://graph.microsoft.com/v1.0/me/messages",
            code=401,
            msg="Unauthorized",
            hdrs=MagicMock(),
            fp=MagicMock(read=MagicMock(return_value=b'{"error":"token_expired"}')),
        )

        # Second call succeeds
        mock_resp_success = MagicMock()
        mock_resp_success.status = 200
        mock_resp_success.headers = {"Content-Type": "application/json"}
        mock_resp_success.read.return_value = b'{"status":"ok"}'
        mock_resp_success.__enter__.return_value = mock_resp_success

        mock_urlopen.side_effect = [err_401, mock_resp_success]

        res = self.client.get(
            self.mock_account, "https://graph.microsoft.com/v1.0/me/messages"
        )
        self.assertEqual(res, {"status": "ok"})
        # Ensure token_manager was called with force_refresh=True
        self.mock_token_mgr.get_access_token.assert_any_call(
            self.mock_account, force_refresh=True
        )
    @patch("urllib.request.urlopen")
    def test_401_retry_succeeds_even_with_zero_max_retries(
        self, mock_urlopen: MagicMock
    ) -> None:
        client_zero_retry = GraphHttpClient(
            self.mock_token_mgr, timeout_ms=5000, max_retries=0
        )
        err_401 = HTTPError(
            url="https://graph.microsoft.com/v1.0/me/messages",
            code=401,
            msg="Unauthorized",
            hdrs=MagicMock(),
            fp=MagicMock(read=MagicMock(return_value=b'{"error":"expired"}')),
        )
        mock_resp_success = MagicMock()
        mock_resp_success.status = 200
        mock_resp_success.headers = {"Content-Type": "application/json"}
        mock_resp_success.read.return_value = b'{"status":"recovered"}'
        mock_resp_success.__enter__.return_value = mock_resp_success

        mock_urlopen.side_effect = [err_401, mock_resp_success]

        res = client_zero_retry.get(
            self.mock_account, "https://graph.microsoft.com/v1.0/me/messages"
        )
        self.assertEqual(res, {"status": "recovered"})


if __name__ == "__main__":
    unittest.main()
