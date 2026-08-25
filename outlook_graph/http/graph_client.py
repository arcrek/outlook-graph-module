"""Resilient HTTP client for interacting with Microsoft Graph API v1.0."""

from __future__ import annotations
import json
import random
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from typing import Any, Literal

from ..auth.token_manager import OutlookTokenManager
from ..types.account import AccountCredentials
from ..types.errors import GraphApiError, RateLimitError


@dataclass(slots=True)
class RequestOptions:
    """Per-request options for GraphHttpClient."""

    headers: dict[str, str] | None = None
    body: Any = None
    search_consistency_eventual: bool = False
    timeout_ms: int | None = None
    max_retries: int | None = None
    initial_retry_delay_ms: int | None = None

    @property
    def searchConsistencyEventual(self) -> bool:
        return self.search_consistency_eventual

    @property
    def timeoutMs(self) -> int | None:
        return self.timeout_ms

    @property
    def maxRetries(self) -> int | None:
        return self.max_retries

    @property
    def initialRetryDelayMs(self) -> int | None:
        return self.initial_retry_delay_ms


class GraphHttpClient:
    """Resilient HTTP client with OAuth token injection, retry, and backoff."""

    def __init__(
        self,
        token_manager: OutlookTokenManager,
        timeout_ms: int = 15000,
        max_retries: int = 3,
        initial_retry_delay_ms: int = 1000,
    ) -> None:
        self.token_manager = token_manager
        self.default_timeout_sec = timeout_ms / 1000.0
        self.default_max_retries = max_retries
        self.default_initial_retry_delay_sec = initial_retry_delay_ms / 1000.0

    def request(
        self,
        account: AccountCredentials,
        url: str,
        method: Literal["GET", "POST", "PATCH", "DELETE"] = "GET",
        options: RequestOptions | None = None,
    ) -> Any:
        """Execute an authenticated HTTP request to Microsoft Graph API."""
        access_token = self.token_manager.get_access_token(account)
        did_retry_auth = False
        attempt = 0

        max_retries = (
            options.max_retries
            if options and options.max_retries is not None
            else self.default_max_retries
        )
        retry_delay_sec = (
            (options.initial_retry_delay_ms / 1000.0)
            if options and options.initial_retry_delay_ms is not None
            else self.default_initial_retry_delay_sec
        )
        timeout_sec = (
            (options.timeout_ms / 1000.0)
            if options and options.timeout_ms is not None
            else self.default_timeout_sec
        )

        while attempt <= max_retries:
            attempt += 1

            headers = {
                "Authorization": f"Bearer {access_token}",
                "Accept": "application/json",
                "User-Agent": "graph-module-python/1.0.0",
            }
            if options and options.headers:
                headers.update(options.headers)

            if options and options.search_consistency_eventual:
                headers["ConsistencyLevel"] = "eventual"

            request_body_bytes: bytes | None = None
            if options and options.body is not None:
                headers["Content-Type"] = "application/json"
                if isinstance(options.body, (dict, list)):
                    request_body_bytes = json.dumps(options.body).encode("utf-8")
                elif isinstance(options.body, str):
                    request_body_bytes = options.body.encode("utf-8")
                elif isinstance(options.body, bytes):
                    request_body_bytes = options.body

            req = urllib.request.Request(
                url,
                data=request_body_bytes,
                headers=headers,
                method=method,
            )

            try:
                with urllib.request.urlopen(req, timeout=timeout_sec) as resp:
                    status_code = resp.status
                    if status_code == 204:
                        return None

                    content_type = resp.headers.get("Content-Type", "")
                    raw_data = resp.read()
                    text_data = raw_data.decode("utf-8", errors="replace")

                    if "application/json" in content_type:
                        try:
                            return json.loads(text_data)
                        except Exception:
                            return text_data
                    return text_data

            except urllib.error.HTTPError as http_err:
                status = http_err.code

                # 401 Unauthorized -> Refresh token and retry once
                if status == 401 and not did_retry_auth:
                    did_retry_auth = True
                    attempt -= 1  # Do not consume normal retry budget for 401 token renewal
                    try:
                        http_err.read()
                        http_err.close()
                    except Exception:
                        pass
                    access_token = self.token_manager.get_access_token(
                        account, force_refresh=True
                    )
                    continue

                # 429 Rate Limit
                if status == 429:
                    retry_header = http_err.headers.get("Retry-After")
                    retry_sec = (
                        float(retry_header)
                        if retry_header and retry_header.isdigit()
                        else retry_delay_sec
                    )
                    retry_ms = int(retry_sec * 1000)

                    try:
                        http_err.read()
                        http_err.close()
                    except Exception:
                        pass

                    if attempt > max_retries:
                        raise RateLimitError(
                            f"Microsoft Graph rate limit reached for account {account.email}. Max retries exceeded.",
                            retry_after_ms=retry_ms,
                            status_code=429,
                        ) from http_err

                    time.sleep(retry_sec + random.uniform(0, 0.2))
                    retry_delay_sec *= 2
                    continue

                # 5xx Server Error
                if 500 <= status < 600:
                    err_text = ""
                    try:
                        err_text = http_err.read().decode("utf-8", errors="replace")
                        http_err.close()
                    except Exception:
                        pass

                    if attempt > max_retries:
                        raise GraphApiError(
                            f"Microsoft Graph server error ({status}): {err_text}",
                            status_code=status,
                            graph_error_code="ServerError",
                        ) from http_err

                    time.sleep(retry_delay_sec + random.uniform(0, 0.2))
                    retry_delay_sec *= 2
                    continue
                # Other HTTP errors (400, 403, 404, etc.)
                err_bytes = http_err.read()
                err_text = err_bytes.decode("utf-8", errors="replace")
                error_code = "UnknownError"
                error_message = http_err.reason
                response_body: Any = None

                content_type = http_err.headers.get("Content-Type", "")
                if "application/json" in content_type:
                    try:
                        data = json.loads(err_text)
                        response_body = data
                        if isinstance(data, dict) and "error" in data:
                            err_obj = data["error"]
                            if isinstance(err_obj, dict):
                                error_code = err_obj.get("code", error_code)
                                error_message = err_obj.get(
                                    "message", error_message
                                )
                    except Exception:
                        pass
                else:
                    error_message = err_text or http_err.reason

                raise GraphApiError(
                    f"Microsoft Graph API error ({status} {error_code}): {error_message}",
                    status_code=status,
                    graph_error_code=error_code,
                    response_body=response_body,
                ) from http_err

            except Exception as err:
                if isinstance(err, (GraphApiError, RateLimitError)):
                    raise err

                if attempt > max_retries:
                    raise GraphApiError(
                        f"Network request failed to '{url}' after {attempt} attempts: {err}"
                    ) from err

                time.sleep(retry_delay_sec + random.uniform(0, 0.2))
                retry_delay_sec *= 2

        raise GraphApiError(
            f"Request to '{url}' failed: max retries exceeded"
        )

    def get(
        self,
        account: AccountCredentials,
        url: str,
        options: RequestOptions | None = None,
    ) -> Any:
        """Execute a GET request."""
        return self.request(account, url, method="GET", options=options)

    def post(
        self,
        account: AccountCredentials,
        url: str,
        body: Any = None,
        options: RequestOptions | None = None,
    ) -> Any:
        """Execute a POST request with payload."""
        req_opts = options or RequestOptions()
        req_opts.body = body
        return self.request(account, url, method="POST", options=req_opts)

    def patch(
        self,
        account: AccountCredentials,
        url: str,
        body: Any = None,
        options: RequestOptions | None = None,
    ) -> Any:
        """Execute a PATCH request with payload."""
        req_opts = options or RequestOptions()
        req_opts.body = body
        return self.request(account, url, method="PATCH", options=req_opts)

    def delete(
        self,
        account: AccountCredentials,
        url: str,
        options: RequestOptions | None = None,
    ) -> Any:
        """Execute a DELETE request."""
        return self.request(account, url, method="DELETE", options=options)
