"""Outlook token manager for OAuth 2.0 token lifecycle and caching."""

from __future__ import annotations
import json
import logging
import random
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from typing import Any

from ..types.account import AccountCredentials
from ..types.auth import (
    TokenCacheEntry,
    TokenManagerOptions,
    TokenResponse,
    TokenRotationCallback,
)
from ..types.errors import RateLimitError, TokenRefreshError

logger = logging.getLogger(__name__)

DEFAULT_SCOPES = [
    "https://graph.microsoft.com/Mail.Read",
    "https://graph.microsoft.com/Mail.ReadWrite",
    "offline_access",
]


class OutlookTokenManager:
    """Manages OAuth 2.0 access token refresh, caching, rotation, and retries."""

    def __init__(self, options: TokenManagerOptions | None = None) -> None:
        self.default_authority = (
            options.authority if options and options.authority else "common"
        )
        self.default_scopes = (
            options.scopes
            if options and options.scopes is not None
            else list(DEFAULT_SCOPES)
        )
        self.expiry_buffer_ms = (
            (options.expiry_buffer_sec if options else 300) * 1000
        )
        self.max_retries = options.max_retries if options else 3
        self.initial_retry_delay_ms = (
            options.initial_retry_delay_ms if options else 1000
        )
        self.request_timeout_sec = (
            (options.request_timeout_ms if options else 15000) / 1000.0
        )
        self.on_token_rotated: TokenRotationCallback | None = (
            options.on_token_rotated if options else None
        )

        self._token_cache: dict[str, TokenCacheEntry] = {}
        self._lock = threading.RLock()
        self._in_flight_locks: dict[str, threading.Lock] = {}

    def get_cached_token(self, email: str) -> TokenCacheEntry | None:
        """Retrieve token cache entry for a given email if available."""
        with self._lock:
            return self._token_cache.get(email.lower())

    def clear_cache(self, email: str | None = None) -> None:
        """Clear cached tokens for a specific email or all accounts."""
        with self._lock:
            if email:
                self._token_cache.pop(email.lower(), None)
            else:
                self._token_cache.clear()

    def is_token_valid(self, entry: TokenCacheEntry) -> bool:
        """Check if a token is still valid with the configured expiry buffer."""
        now_ms = int(time.time() * 1000)
        return now_ms < (entry.expires_at - self.expiry_buffer_ms)

    def get_access_token(
        self, account: AccountCredentials, force_refresh: bool = False
    ) -> str:
        """Get a valid access token from cache or refresh proactively with de-duplication."""
        email_key = account.email.lower()
        if not force_refresh:
            cached = self.get_cached_token(email_key)
            if cached and self.is_token_valid(cached):
                return cached.access_token

        with self._lock:
            if email_key not in self._in_flight_locks:
                self._in_flight_locks[email_key] = threading.Lock()
            flight_lock = self._in_flight_locks[email_key]

        with flight_lock:
            if not force_refresh:
                cached = self.get_cached_token(email_key)
                if cached and self.is_token_valid(cached):
                    return cached.access_token
            response = self._execute_refresh(account)
            return response.access_token
    def refresh_access_token(self, account: AccountCredentials) -> TokenResponse:
        """Refresh the OAuth 2.0 access token with de-duplication."""
        email_key = account.email.lower()

        with self._lock:
            if email_key not in self._in_flight_locks:
                self._in_flight_locks[email_key] = threading.Lock()
            flight_lock = self._in_flight_locks[email_key]

        with flight_lock:
            return self._execute_refresh(account)

    def _execute_refresh(self, account: AccountCredentials) -> TokenResponse:
        authority = account.authority or self.default_authority
        token_url = f"https://login.microsoftonline.com/{authority}/oauth2/v2.0/token"
        scope = " ".join(self.default_scopes)

        data = {
            "client_id": account.client_id,
            "grant_type": "refresh_token",
            "refresh_token": account.refresh_token,
            "scope": scope,
        }
        encoded_data = urllib.parse.urlencode(data).encode("utf-8")

        headers = {
            "Content-Type": "application/x-www-form-urlencoded",
            "Accept": "application/json",
            "User-Agent": "graph-module-python/1.0.0",
        }

        attempt = 0
        delay_sec = self.initial_retry_delay_ms / 1000.0

        while attempt <= self.max_retries:
            attempt += 1
            req = urllib.request.Request(
                token_url, data=encoded_data, headers=headers, method="POST"
            )

            try:
                with urllib.request.urlopen(
                    req, timeout=self.request_timeout_sec
                ) as resp:
                    status_code = resp.status
                    resp_body = resp.read().decode("utf-8")
                    json_data: dict[str, Any] = json.loads(resp_body)

                token_response = TokenResponse(
                    token_type=json_data.get("token_type", "Bearer"),
                    scope=json_data.get("scope", scope),
                    expires_in=int(json_data.get("expires_in", 3600)),
                    ext_expires_in=json_data.get("ext_expires_in"),
                    access_token=json_data["access_token"],
                    refresh_token=json_data.get("refresh_token"),
                    id_token=json_data.get("id_token"),
                )

                old_refresh_token = account.refresh_token
                if (
                    token_response.refresh_token
                    and token_response.refresh_token != old_refresh_token
                ):
                    account.refresh_token = token_response.refresh_token
                    if self.on_token_rotated:
                        try:
                            self.on_token_rotated(
                                old_refresh_token,
                                token_response.refresh_token,
                                account,
                            )
                        except Exception as listener_err:
                            logger.warning(
                                "on_token_rotated callback failed for %s: %s",
                                account.email,
                                listener_err,
                            )

                expires_at = (
                    int(time.time() * 1000) + token_response.expires_in * 1000
                )
                with self._lock:
                    self._token_cache[account.email.lower()] = TokenCacheEntry(
                        access_token=token_response.access_token,
                        refresh_token=token_response.refresh_token
                        or account.refresh_token,
                        expires_at=expires_at,
                        scope=token_response.scope,
                    )

                return token_response

            except urllib.error.HTTPError as http_err:
                status = http_err.code
                error_body = http_err.read().decode("utf-8", errors="replace")
                parsed_json: dict[str, Any] = {}
                try:
                    parsed_json = json.loads(error_body)
                except Exception:
                    pass

                if status == 429:
                    retry_header = http_err.headers.get("Retry-After")
                    retry_sec = (
                        float(retry_header)
                        if retry_header and retry_header.isdigit()
                        else delay_sec
                    )
                    retry_ms = int(retry_sec * 1000)

                    if attempt > self.max_retries:
                        raise RateLimitError(
                            f"Rate limited (HTTP 429) refreshing token for {account.email}. Max retries exceeded.",
                            retry_after_ms=retry_ms,
                            status_code=429,
                        ) from http_err

                    time.sleep(retry_sec + random.uniform(0, 0.2))
                    delay_sec *= 2
                    continue

                if 500 <= status < 600:
                    if attempt > self.max_retries:
                        raise TokenRefreshError(
                            f"Server error (HTTP {status}) while refreshing token for {account.email}: {error_body}",
                            status_code=status,
                        ) from http_err

                    time.sleep(delay_sec + random.uniform(0, 0.2))
                    delay_sec *= 2
                    continue

                # 400 Bad Request / 401 Unauthorized etc.
                error_desc = (
                    parsed_json.get("error_description")
                    or parsed_json.get("error")
                    or http_err.reason
                )
                raise TokenRefreshError(
                    f"Failed to refresh OAuth token for account {account.email}: {error_desc}",
                    status_code=status,
                    error_response=parsed_json,
                ) from http_err

            except Exception as err:
                if isinstance(err, (TokenRefreshError, RateLimitError)):
                    raise err

                if attempt > self.max_retries:
                    raise TokenRefreshError(
                        f"Network error refreshing token for {account.email} after {attempt} attempts: {err}"
                    ) from err

                time.sleep(delay_sec + random.uniform(0, 0.2))
                delay_sec *= 2

        raise TokenRefreshError(
            f"Failed to refresh token for {account.email}: max retries exceeded"
        )

    # TypeScript naming parity aliases
    getCachedToken = get_cached_token
    clearCache = clear_cache
    isTokenValid = is_token_valid
    getAccessToken = get_access_token
    refreshAccessToken = refresh_access_token
