"""Account credentials and parsing options data models."""

from __future__ import annotations
from dataclasses import dataclass


@dataclass(slots=True)
class AccountCredentials:
    """Represents a set of Microsoft account credentials."""

    email: str
    refresh_token: str
    client_id: str
    password: str | None = None
    authority: str | None = None
    recovery_email: str | None = None

    @property
    def refreshToken(self) -> str:
        """Alias for refresh_token to maintain TypeScript naming parity."""
        return self.refresh_token

    @property
    def clientId(self) -> str:
        """Alias for client_id to maintain TypeScript naming parity."""
        return self.client_id

    @property
    def recoveryEmail(self) -> str | None:
        """Alias for recovery_email to maintain TypeScript naming parity."""
        return self.recovery_email

@dataclass(slots=True)
class ParseAccountOptions:
    """Options for parsing account strings or files."""

    trim: bool = True
    ignore_empty_lines: bool = True
    ignore_comments: bool = True
    default_authority: str | None = None

    @property
    def ignoreEmptyLines(self) -> bool:
        return self.ignore_empty_lines

    @property
    def ignoreComments(self) -> bool:
        return self.ignore_comments

    @property
    def defaultAuthority(self) -> str | None:
        return self.default_authority
