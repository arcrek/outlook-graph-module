"""Mailbox messages, attachments, query options, and polling models."""

from __future__ import annotations
from dataclasses import dataclass, field
from datetime import datetime
from typing import Callable, Any, Literal


@dataclass(slots=True)
class EmailAddress:
    """Email address structure representing a contact or mailbox."""

    address: str
    name: str | None = None


@dataclass(slots=True)
class EmailRecipient:
    """Recipient wrapper with an inner emailAddress field."""

    email_address: EmailAddress

    @property
    def emailAddress(self) -> EmailAddress:
        return self.email_address


@dataclass(slots=True)
class ItemBody:
    """Email body content with contentType ('text' or 'html')."""

    content_type: Literal["text", "html"]
    content: str

    @property
    def contentType(self) -> Literal["text", "html"]:
        return self.content_type


@dataclass(slots=True)
class OutlookAttachment:
    """Email attachment metadata and binary content."""

    id: str
    name: str
    content_type: str
    size: int
    is_inline: bool = False
    content_bytes: str | None = None
    data: bytes | None = None

    @property
    def contentType(self) -> str:
        return self.content_type

    @property
    def isInline(self) -> bool:
        return self.is_inline

    @property
    def contentBytes(self) -> str | None:
        return self.content_bytes


@dataclass(slots=True)
class OutlookMessage:
    """Full Outlook email message model."""

    id: str
    subject: str
    body: ItemBody
    received_date_time: str
    has_attachments: bool
    is_read: bool
    body_preview: str | None = None
    conversation_id: str | None = None
    from_recipient: EmailRecipient | None = None
    to_recipients: list[EmailRecipient] = field(default_factory=list)
    cc_recipients: list[EmailRecipient] = field(default_factory=list)
    sent_date_time: str | None = None
    web_link: str | None = None
    attachments: list[OutlookAttachment] | None = None
    raw: dict[str, Any] | None = None

    @property
    def receivedDateTime(self) -> str:
        return self.received_date_time

    @property
    def bodyPreview(self) -> str | None:
        return self.body_preview

    @property
    def conversationId(self) -> str | None:
        return self.conversation_id

    @property
    def from_(self) -> EmailRecipient | None:
        return self.from_recipient

    @property
    def toRecipients(self) -> list[EmailRecipient]:
        return self.to_recipients

    @property
    def ccRecipients(self) -> list[EmailRecipient]:
        return self.cc_recipients

    @property
    def sentDateTime(self) -> str | None:
        return self.sent_date_time

    @property
    def hasAttachments(self) -> bool:
        return self.has_attachments

    @property
    def isRead(self) -> bool:
        return self.is_read

    @property
    def webLink(self) -> str | None:
        return self.web_link


@dataclass(slots=True)
class GetMessagesOptions:
    """Query filter and pagination options for fetching messages."""

    folder: str | None = None
    top: int | None = None
    skip: int | None = None
    unread_only: bool | None = None
    subject_contains: str | None = None
    from_contains: str | None = None
    received_after: datetime | str | None = None
    search: str | None = None
    order_by: str | None = None
    select: list[str] | None = None
    filter: str | None = None

    @property
    def unreadOnly(self) -> bool | None:
        return self.unread_only

    @property
    def subjectContains(self) -> str | None:
        return self.subject_contains

    @property
    def fromContains(self) -> str | None:
        return self.from_contains

    @property
    def receivedAfter(self) -> datetime | str | None:
        return self.received_after

    @property
    def orderBy(self) -> str | None:
        return self.order_by


@dataclass(slots=True)
class WaitForEmailOptions:
    """Options for non-blocking email polling loop."""

    subject_contains: str | None = None
    from_contains: str | None = None
    received_after: datetime | str | None = None
    folder: str | None = None
    unread_only: bool | None = None
    timeout_ms: int = 60000
    interval_ms: int = 3000
    predicate: Callable[[OutlookMessage], bool] | None = None
    mark_as_read_after_match: bool = False

    @property
    def subjectContains(self) -> str | None:
        return self.subject_contains

    @property
    def fromContains(self) -> str | None:
        return self.from_contains

    @property
    def receivedAfter(self) -> datetime | str | None:
        return self.received_after

    @property
    def unreadOnly(self) -> bool | None:
        return self.unread_only

    @property
    def timeoutMs(self) -> int:
        return self.timeout_ms

    @property
    def intervalMs(self) -> int:
        return self.interval_ms

    @property
    def markAsReadAfterMatch(self) -> bool:
        return self.mark_as_read_after_match
