"""Microsoft Outlook Graph API Python Module.

A high-performance Python module for interacting with Microsoft Outlook and Hotmail
mailboxes via Microsoft Graph API v1.0, with OAuth2 token rotation, rate-limit resilience,
OData query building, email polling, and multi-language (EN/VI) OTP extraction.
"""

from .types import (
    AccountCredentials,
    ParseAccountOptions,
    TokenResponse,
    TokenCacheEntry,
    TokenRotationCallback,
    TokenManagerOptions,
    EmailAddress,
    EmailRecipient,
    ItemBody,
    OutlookAttachment,
    OutlookMessage,
    GetMessagesOptions,
    WaitForEmailOptions,
    OtpExtractOptions,
    OtpResult,
    GraphModuleError,
    AccountParseError,
    TokenRefreshError,
    GraphApiError,
    RateLimitError,
    TimeoutError,
)
from .parser import (
    parse_account_line,
    parse_account_lines,
    parse_account_file,
    parseAccountLine,
    parseAccountLines,
    parseAccountFile,
)
from .auth import OutlookTokenManager
from .http import (
    GraphHttpClient,
    RequestOptions,
    build_odata_query,
    escape_odata_string,
    get_mail_endpoint,
    buildODataQuery,
    escapeODataString,
    getMailEndpoint,
)
from .mail import OutlookMailClient
from .otp import (
    clean_html,
    extract_otp,
    extract_all_otps,
    cleanHtml,
    extractOtp,
    extractAllOtps,
)

__version__ = "1.0.0"

__all__ = [
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
    "GraphModuleError",
    "AccountParseError",
    "TokenRefreshError",
    "GraphApiError",
    "RateLimitError",
    "TimeoutError",
    "parse_account_line",
    "parse_account_lines",
    "parse_account_file",
    "parseAccountLine",
    "parseAccountLines",
    "parseAccountFile",
    "OutlookTokenManager",
    "GraphHttpClient",
    "RequestOptions",
    "build_odata_query",
    "escape_odata_string",
    "get_mail_endpoint",
    "buildODataQuery",
    "escapeODataString",
    "getMailEndpoint",
    "OutlookMailClient",
    "clean_html",
    "extract_otp",
    "extract_all_otps",
    "cleanHtml",
    "extractOtp",
    "extractAllOtps",
]
