"""High-level Outlook Mail client for message operations, attachments, and polling."""

from __future__ import annotations
import base64
import time
from datetime import datetime, timedelta, timezone
from urllib.parse import quote
from typing import Any

from ..auth.token_manager import OutlookTokenManager
from ..http.graph_client import GraphHttpClient, RequestOptions
from ..http.query_builder import build_odata_query, get_mail_endpoint
from ..otp.otp_extractor import extract_otp
from ..types.account import AccountCredentials
from ..types.auth import TokenManagerOptions
from ..types.errors import TimeoutError, TokenRefreshError, GraphApiError
from ..types.mail import (
    EmailAddress,
    EmailRecipient,
    GetMessagesOptions,
    ItemBody,
    OutlookAttachment,
    OutlookMessage,
    WaitForEmailOptions,
)
from ..types.otp import OtpExtractOptions, OtpResult


class OutlookMailClient:
    """Client for performing mailbox operations via Microsoft Graph API v1.0."""

    def __init__(
        self,
        token_manager_or_options: OutlookTokenManager
        | TokenManagerOptions
        | None = None,
    ) -> None:
        if isinstance(token_manager_or_options, OutlookTokenManager):
            self.token_manager = token_manager_or_options
        else:
            self.token_manager = OutlookTokenManager(token_manager_or_options)
        self.http_client = GraphHttpClient(self.token_manager)

    @property
    def httpClient(self) -> GraphHttpClient:
        return self.http_client

    def get_messages(
        self,
        account: AccountCredentials,
        options: GetMessagesOptions | None = None,
    ) -> list[OutlookMessage]:
        """Fetch messages with OData filtering, folder selection, and sorting."""
        folder = options.folder if options else None
        endpoint = get_mail_endpoint(folder)
        query_string = build_odata_query(options)
        url = f"{endpoint}{query_string}"

        req_options = RequestOptions(
            search_consistency_eventual=bool(options and options.search)
        )
        res = self.http_client.get(account, url, options=req_options)

        raw_list: list[dict[str, Any]] = (
            res.get("value", []) if isinstance(res, dict) else []
        )
        return [self._map_to_outlook_message(raw) for raw in raw_list]

    def get_latest_message(
        self,
        account: AccountCredentials,
        options: GetMessagesOptions | None = None,
    ) -> OutlookMessage | None:
        """Get the single most recent message matching criteria."""
        merged_options = GetMessagesOptions(
            folder=options.folder if options else None,
            top=1,
            skip=options.skip if options else None,
            unread_only=options.unread_only if options else None,
            subject_contains=options.subject_contains if options else None,
            from_contains=options.from_contains if options else None,
            received_after=options.received_after if options else None,
            search=options.search if options else None,
            order_by=options.order_by if options else None,
            select=options.select if options else None,
            filter=options.filter if options else None,
        )
        messages = self.get_messages(account, merged_options)
        return messages[0] if messages else None

    def get_message_by_id(
        self,
        account: AccountCredentials,
        message_id: str,
        options: dict[str, Any] | None = None,
    ) -> OutlookMessage:
        """Fetch a single message by ID, including its complete body content."""
        select_fields = options.get("select") if options else None
        select_query = (
            f"?$select={','.join(select_fields)}" if select_fields else ""
        )
        url = f"https://graph.microsoft.com/v1.0/me/messages/{quote(message_id)}{select_query}"

        raw = self.http_client.get(account, url)
        return self._map_to_outlook_message(raw)

    def get_message_attachments(
        self,
        account: AccountCredentials,
        message_id: str,
    ) -> list[OutlookAttachment]:
        """Fetch all attachments for a specific message ID and decode base64 content."""
        url = f"https://graph.microsoft.com/v1.0/me/messages/{quote(message_id)}/attachments"
        res = self.http_client.get(account, url)

        raw_list: list[dict[str, Any]] = (
            res.get("value", []) if isinstance(res, dict) else []
        )
        attachments: list[OutlookAttachment] = []

        for item in raw_list:
            content_bytes_str = item.get("contentBytes")
            data_bytes: bytes | None = None
            if content_bytes_str:
                try:
                    data_bytes = base64.b64decode(content_bytes_str)
                except Exception:
                    pass

            attachments.append(
                OutlookAttachment(
                    id=item.get("id", ""),
                    name=item.get("name", ""),
                    content_type=item.get(
                        "contentType", "application/octet-stream"
                    ),
                    size=int(item.get("size", 0)),
                    is_inline=bool(item.get("isInline", False)),
                    content_bytes=content_bytes_str,
                    data=data_bytes,
                )
            )

        return attachments

    def mark_as_read(
        self,
        account: AccountCredentials,
        message_id: str,
        is_read: bool = True,
    ) -> None:
        """Update read status of a message."""
        url = f"https://graph.microsoft.com/v1.0/me/messages/{quote(message_id)}"
        self.http_client.patch(account, url, body={"isRead": is_read})

    def delete_message(
        self,
        account: AccountCredentials,
        message_id: str,
    ) -> None:
        """Permanently delete a message."""
        url = f"https://graph.microsoft.com/v1.0/me/messages/{quote(message_id)}"
        self.http_client.delete(account, url)

    def wait_for_email(
        self,
        account: AccountCredentials,
        options: WaitForEmailOptions | None = None,
    ) -> OutlookMessage:
        """Poll the mailbox until a message matching filter/predicate arrives."""
        timeout_ms = options.timeout_ms if options else 60000
        interval_sec = (options.interval_ms if options else 3000) / 1000.0
        received_after = (
            options.received_after
            if options and options.received_after is not None
            else (datetime.now(timezone.utc) - timedelta(seconds=60))
        )

        start_time = time.time()
        timeout_sec = timeout_ms / 1000.0

        while True:
            elapsed = time.time() - start_time
            if elapsed > timeout_sec:
                raise TimeoutError(
                    f"Timed out after {timeout_ms}ms waiting for matching email on account {account.email}"
                )

            try:
                query_opts = GetMessagesOptions(
                    folder=options.folder if options else None,
                    unread_only=options.unread_only if options else None,
                    subject_contains=options.subject_contains
                    if options
                    else None,
                    from_contains=options.from_contains if options else None,
                    received_after=received_after,
                    top=10,
                    order_by="receivedDateTime desc",
                )
                messages = self.get_messages(account, query_opts)

                for msg in messages:
                    matched = True

                    if options and options.subject_contains:
                        if (
                            options.subject_contains.lower()
                            not in msg.subject.lower()
                        ):
                            matched = False

                    if matched and options and options.from_contains:
                        sender_addr = (
                            msg.from_recipient.email_address.address.lower()
                            if msg.from_recipient
                            and msg.from_recipient.email_address
                            else ""
                        )
                        if options.from_contains.lower() not in sender_addr:
                            matched = False

                    if matched and options and options.predicate:
                        matched = options.predicate(msg)

                    if matched:
                        if options and options.mark_as_read_after_match:
                            try:
                                self.mark_as_read(account, msg.id, True)
                            except Exception:
                                pass
                        return msg

            except Exception as err:
                if isinstance(err, TimeoutError):
                    raise err
                if isinstance(err, (TokenRefreshError, GraphApiError)):
                    raise err
                # RateLimitError (already retried internally) and any other
                # unexpected error: keep polling until timeout.

            time.sleep(interval_sec)

    def extract_otp(
        self,
        input_data: OutlookMessage | str,
        options: OtpExtractOptions | None = None,
    ) -> OtpResult | None:
        """Extract verification code from an OutlookMessage or plain text."""
        return extract_otp(input_data, options)

    def _map_to_outlook_message(
        self, raw: dict[str, Any]
    ) -> OutlookMessage:
        from_raw = raw.get("from")
        from_recipient: EmailRecipient | None = None
        if isinstance(from_raw, dict) and "emailAddress" in from_raw:
            addr_data = from_raw["emailAddress"]
            if isinstance(addr_data, dict):
                from_recipient = EmailRecipient(
                    email_address=EmailAddress(
                        name=addr_data.get("name"),
                        address=addr_data.get("address", ""),
                    )
                )

        body_raw = raw.get("body")
        if isinstance(body_raw, dict):
            body = ItemBody(
                content_type="html"
                if body_raw.get("contentType") == "html"
                else "text",
                content=body_raw.get("content", ""),
            )
        else:
            body = ItemBody(content_type="text", content="")

        to_recipients: list[EmailRecipient] = []
        for r in raw.get("toRecipients", []):
            if isinstance(r, dict) and isinstance(r.get("emailAddress"), dict):
                to_recipients.append(
                    EmailRecipient(
                        email_address=EmailAddress(
                            name=r["emailAddress"].get("name"),
                            address=r["emailAddress"].get("address", ""),
                        )
                    )
                )

        cc_recipients: list[EmailRecipient] = []
        for r in raw.get("ccRecipients", []):
            if isinstance(r, dict) and isinstance(r.get("emailAddress"), dict):
                cc_recipients.append(
                    EmailRecipient(
                        email_address=EmailAddress(
                            name=r["emailAddress"].get("name"),
                            address=r["emailAddress"].get("address", ""),
                        )
                    )
                )

        return OutlookMessage(
            id=raw.get("id", ""),
            conversation_id=raw.get("conversationId"),
            subject=raw.get("subject") or "(No Subject)",
            body_preview=raw.get("bodyPreview"),
            body=body,
            from_recipient=from_recipient,
            to_recipients=to_recipients,
            cc_recipients=cc_recipients,
            received_date_time=raw.get(
                "receivedDateTime", datetime.now(timezone.utc).isoformat()
            ),
            sent_date_time=raw.get("sentDateTime"),
            has_attachments=bool(raw.get("hasAttachments", False)),
            is_read=bool(raw.get("isRead", False)),
            web_link=raw.get("webLink"),
            raw=raw,
        )

    # TypeScript naming parity aliases
    getMessages = get_messages
    getLatestMessage = get_latest_message
    getMessageById = get_message_by_id
    getMessageAttachments = get_message_attachments
    markAsRead = mark_as_read
    deleteMessage = delete_message
    waitForEmail = wait_for_email
    extractOtp = extract_otp
