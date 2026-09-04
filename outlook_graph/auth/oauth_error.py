"""OAuth error diagnostic parser for Microsoft Azure AD / Entra ID responses."""

from __future__ import annotations
import re
from ..types.errors import TokenRefreshError, TokenErrorReason


def parse_oauth_error(
    status_code: int,
    data: dict[str, object] | None,
    email: str,
) -> TokenRefreshError:
    """Parse OAuth error response and build a diagnostics-rich TokenRefreshError."""
    data_dict: dict[str, object] = data or {}
    raw_desc = data_dict.get("error_description") or data_dict.get("error") or "Bad Request"
    error_desc = str(raw_desc)

    aadsts_code: int | None = None

    # Check error_codes list
    error_codes = data_dict.get("error_codes")
    if isinstance(error_codes, list) and len(error_codes) > 0:
        first_code: object = error_codes[0]
        if isinstance(first_code, int):
            aadsts_code = first_code
        elif isinstance(first_code, str) and first_code.isdigit():
            aadsts_code = int(first_code)

    # Fallback to regex in error description
    if aadsts_code is None:
        match = re.search(r"AADSTS(\d+)", error_desc, re.IGNORECASE)
        if match:
            try:
                aadsts_code = int(match.group(1))
            except ValueError:
                pass

    diagnostic_reason: TokenErrorReason = "UNKNOWN"
    remediation = (
        "Token refresh failed. Verify account credentials and Azure AD configuration."
    )

    if aadsts_code == 70000:
        if re.search(r"different client", error_desc, re.IGNORECASE):
            diagnostic_reason = "CLIENT_ID_MISMATCH"
            remediation = (
                "The refresh token was issued for a different client_id. "
                "Verify that client_id matches the application that generated the token."
            )
        else:
            diagnostic_reason = "EXPIRED_OR_REVOKED_GRANT"
            remediation = (
                "The refresh token is expired or revoked. "
                "Re-authenticate to obtain a new refresh token."
            )
    elif aadsts_code in (700082, 70008, 700084, 50173):
        diagnostic_reason = "EXPIRED_OR_REVOKED_GRANT"
        remediation = (
            "The refresh token has expired or was revoked. "
            "Re-authenticate to obtain a new refresh token."
        )
    elif aadsts_code == 700016:
        diagnostic_reason = "INVALID_CLIENT"
        remediation = (
            "The client_id was not found in Azure AD. "
            "Verify your client_id configuration."
        )
    elif aadsts_code == 50053:
        diagnostic_reason = "ACCOUNT_LOCKED"
        remediation = (
            "The account is locked due to multiple failed sign-in attempts. "
            "Unlock it via Microsoft account security."
        )
    elif aadsts_code == 50057:
        diagnostic_reason = "ACCOUNT_DISABLED"
        remediation = (
            "The account is disabled. Contact Microsoft support or your administrator."
        )
    elif aadsts_code in (50076, 50079, 53003):
        diagnostic_reason = "INTERACTION_REQUIRED"
        remediation = (
            "Interactive sign-in or multi-factor authentication (MFA) is required by security policy."
        )
    elif aadsts_code == 70011:
        diagnostic_reason = "INVALID_SCOPE"
        remediation = (
            "The requested OAuth scopes are invalid for this resource. Check requested OAuth scopes."
        )
    elif data_dict.get("error") == "invalid_grant":
        diagnostic_reason = "EXPIRED_OR_REVOKED_GRANT"
        remediation = (
            "The grant or refresh token is invalid, expired, or revoked. "
            "Re-authenticate to obtain a new token."
        )

    message = f"Failed to refresh OAuth token for account {email}: {error_desc}"
    return TokenRefreshError(
        message=message,
        status_code=status_code,
        error_response=dict(data_dict),
        aadsts_code=aadsts_code,
        diagnostic_reason=diagnostic_reason,
        remediation=remediation,
    )
