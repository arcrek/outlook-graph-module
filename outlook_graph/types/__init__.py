"""Types package export."""

from .errors import (
    GraphModuleError,
    AccountParseError,
    TokenRefreshError,
    GraphApiError,
    RateLimitError,
    TimeoutError,
)
from .account import AccountCredentials, ParseAccountOptions
from .auth import (
    TokenResponse,
    TokenCacheEntry,
    TokenRotationCallback,
    TokenManagerOptions,
)
from .mail import (
    EmailAddress,
    EmailRecipient,
    ItemBody,
    OutlookAttachment,
    OutlookMessage,
    GetMessagesOptions,
    WaitForEmailOptions,
)
from .otp import OtpExtractOptions, OtpResult

__all__ = [
    "GraphModuleError",
    "AccountParseError",
    "TokenRefreshError",
    "GraphApiError",
    "RateLimitError",
    "TimeoutError",
    "AccountCredentials",
    "ParseAccountOptions",
    "TokenResponse",
    "TokenCacheEntry",
    "TokenRotationCallback",
    "TokenManagerOptions",
    "EmailAddress",
    "EmailRecipient",
    "ItemBody",
    "OutlookAttachment",
    "OutlookMessage",
    "GetMessagesOptions",
    "WaitForEmailOptions",
    "OtpExtractOptions",
    "OtpResult",
]
