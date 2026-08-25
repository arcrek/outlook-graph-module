"""Authentication data models and options."""

from __future__ import annotations
from dataclasses import dataclass
from typing import Callable, Any
from .account import AccountCredentials

TokenRotationCallback = Callable[[str, str, AccountCredentials], Any]


@dataclass(slots=True)
class TokenResponse:
    """OAuth 2.0 token response from Microsoft Identity endpoint."""

    token_type: str
    scope: str
    expires_in: int
    access_token: str
    ext_expires_in: int | None = None
    refresh_token: str | None = None
    id_token: str | None = None


@dataclass(slots=True)
class TokenCacheEntry:
    """Cached access token with expiration and scope metadata."""

    access_token: str
    refresh_token: str
    expires_at: int  # Epoch millisecond
    scope: str

    @property
    def accessToken(self) -> str:
        return self.access_token

    @property
    def refreshToken(self) -> str:
        return self.refresh_token

    @property
    def expiresAt(self) -> int:
        return self.expires_at


@dataclass(slots=True)
class TokenManagerOptions:
    """Configuration options for OutlookTokenManager."""

    authority: str = "common"
    scopes: list[str] | None = None
    expiry_buffer_sec: int = 300
    max_retries: int = 3
    initial_retry_delay_ms: int = 1000
    request_timeout_ms: int = 15000
    on_token_rotated: TokenRotationCallback | None = None
