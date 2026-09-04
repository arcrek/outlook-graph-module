"""Domain error hierarchy for the Microsoft Outlook Graph API module."""

from typing import Any, Literal

TokenErrorReason = Literal[
    "EXPIRED_OR_REVOKED_GRANT",
    "CLIENT_ID_MISMATCH",
    "INVALID_CLIENT",
    "ACCOUNT_LOCKED",
    "ACCOUNT_DISABLED",
    "INTERACTION_REQUIRED",
    "INVALID_SCOPE",
    "UNKNOWN",
]


class GraphModuleError(Exception):
    """Base exception for all Graph module errors."""

    def __init__(self, message: str, code: str | None = None) -> None:
        super().__init__(message)
        self.message = message
        self.code = code

    def __repr__(self) -> str:
        if self.code:
            return f"{self.__class__.__name__}(code={self.code!r}, message={self.message!r})"
        return f"{self.__class__.__name__}(message={self.message!r})"


class AccountParseError(GraphModuleError):
    """Raised when an account credential line or file cannot be parsed."""

    def __init__(
        self,
        message: str,
        raw_line: str | None = None,
        line_number: int | None = None,
    ) -> None:
        super().__init__(message, "ACCOUNT_PARSE_ERROR")
        self.raw_line = raw_line
        self.line_number = line_number


class TokenRefreshError(GraphModuleError):
    """Raised when an OAuth 2.0 token refresh operation fails."""

    def __init__(
        self,
        message: str,
        status_code: int | None = None,
        error_response: dict[str, Any] | None = None,
        aadsts_code: int | None = None,
        diagnostic_reason: TokenErrorReason | str | None = None,
        remediation: str | None = None,
    ) -> None:
        super().__init__(message, "TOKEN_REFRESH_ERROR")
        self.status_code = status_code
        self.error_response = error_response
        self.aadsts_code = aadsts_code
        self.diagnostic_reason = diagnostic_reason
        self.remediation = remediation

    @property
    def aadstsCode(self) -> int | None:
        """TypeScript camelCase alias for aadsts_code."""
        return self.aadsts_code

    @property
    def diagnosticReason(self) -> TokenErrorReason | str | None:
        """TypeScript camelCase alias for diagnostic_reason."""
        return self.diagnostic_reason

class GraphApiError(GraphModuleError):
    """Raised when the Microsoft Graph API returns an HTTP error response."""

    def __init__(
        self,
        message: str,
        status_code: int | None = None,
        graph_error_code: str | None = None,
        response_body: Any = None,
    ) -> None:
        super().__init__(message, "GRAPH_API_ERROR")
        self.status_code = status_code
        self.graph_error_code = graph_error_code
        self.response_body = response_body


class RateLimitError(GraphModuleError):
    """Raised when Microsoft Graph API throttles the request (HTTP 429)."""

    def __init__(
        self,
        message: str,
        retry_after_ms: int,
        status_code: int = 429,
    ) -> None:
        super().__init__(message, "RATE_LIMIT_ERROR")
        self.retry_after_ms = retry_after_ms
        self.status_code = status_code


class TimeoutError(GraphModuleError, TimeoutError):
    """Raised when a polling or waiting operation times out."""

    def __init__(self, message: str) -> None:
        super().__init__(message, "TIMEOUT_ERROR")
