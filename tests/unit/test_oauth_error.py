"""Unit tests for OAuth error diagnostics."""

import unittest
from outlook_graph.auth.oauth_error import parse_oauth_error
from outlook_graph.types.errors import TokenRefreshError


class TestOAuthErrorDiagnostics(unittest.TestCase):
    def test_classify_aadsts70000_expired(self) -> None:
        err = parse_oauth_error(
            400,
            {
                "error": "invalid_grant",
                "error_description": "AADSTS70000: The user could not be authenticated as the grant is expired. The user must sign in again.",
                "error_codes": [70000],
            },
            "test@hotmail.com",
        )
        self.assertIsInstance(err, TokenRefreshError)
        self.assertEqual(err.status_code, 400)
        self.assertEqual(err.aadsts_code, 70000)
        self.assertEqual(err.aadstsCode, 70000)
        self.assertEqual(err.diagnostic_reason, "EXPIRED_OR_REVOKED_GRANT")
        self.assertEqual(err.diagnosticReason, "EXPIRED_OR_REVOKED_GRANT")
        self.assertIn("expired or revoked", err.remediation or "")

    def test_classify_aadsts70000_different_client(self) -> None:
        err = parse_oauth_error(
            400,
            {
                "error": "invalid_grant",
                "error_description": "AADSTS70000: The token was issued for a different client id",
            },
            "test@hotmail.com",
        )
        self.assertEqual(err.aadsts_code, 70000)
        self.assertEqual(err.diagnostic_reason, "CLIENT_ID_MISMATCH")
        self.assertIn("different client_id", err.remediation or "")

    def test_classify_aadsts700082(self) -> None:
        err = parse_oauth_error(
            400,
            {
                "error": "invalid_grant",
                "error_description": "AADSTS700082: The refresh token has expired",
            },
            "test@hotmail.com",
        )
        self.assertEqual(err.aadsts_code, 700082)
        self.assertEqual(err.diagnostic_reason, "EXPIRED_OR_REVOKED_GRANT")

    def test_classify_aadsts700016(self) -> None:
        err = parse_oauth_error(
            400,
            {
                "error": "invalid_client",
                "error_description": "AADSTS700016: Application was not found",
                "error_codes": [700016],
            },
            "test@hotmail.com",
        )
        self.assertEqual(err.aadsts_code, 700016)
        self.assertEqual(err.diagnostic_reason, "INVALID_CLIENT")

    def test_classify_aadsts50053_locked(self) -> None:
        err = parse_oauth_error(
            400,
            {
                "error": "invalid_grant",
                "error_description": "AADSTS50053: The account is locked",
            },
            "test@hotmail.com",
        )
        self.assertEqual(err.aadsts_code, 50053)
        self.assertEqual(err.diagnostic_reason, "ACCOUNT_LOCKED")

    def test_classify_aadsts50076_mfa(self) -> None:
        err = parse_oauth_error(
            400,
            {
                "error": "invalid_grant",
                "error_description": "AADSTS50076: Due to configuration MFA is required",
            },
            "test@hotmail.com",
        )
        self.assertEqual(err.aadsts_code, 50076)
        self.assertEqual(err.diagnostic_reason, "INTERACTION_REQUIRED")

    def test_classify_aadsts70011_scope(self) -> None:
        err = parse_oauth_error(
            400,
            {
                "error": "invalid_scope",
                "error_description": "AADSTS70011: The scope is invalid",
            },
            "test@hotmail.com",
        )
        self.assertEqual(err.aadsts_code, 70011)
        self.assertEqual(err.diagnostic_reason, "INVALID_SCOPE")

    def test_fallback_invalid_grant_without_code(self) -> None:
        err = parse_oauth_error(
            400,
            {
                "error": "invalid_grant",
                "error_description": "The provided refresh token is expired or revoked.",
            },
            "test@hotmail.com",
        )
        self.assertIsNone(err.aadsts_code)
        self.assertEqual(err.diagnostic_reason, "EXPIRED_OR_REVOKED_GRANT")

    def test_unknown_error(self) -> None:
        err = parse_oauth_error(
            400,
            {
                "error": "server_error",
                "error_description": "Internal server problem",
            },
            "test@hotmail.com",
        )
        self.assertIsNone(err.aadsts_code)
        self.assertEqual(err.diagnostic_reason, "UNKNOWN")


if __name__ == "__main__":
    _ = unittest.main()
