"""OTP extraction options and result data models."""

from __future__ import annotations
from dataclasses import dataclass
from typing import Pattern, Literal


@dataclass(slots=True)
class OtpExtractOptions:
    """Configuration options for extracting OTP codes."""

    preferred_length: int | None = None
    allow_alphanumeric: bool = True
    custom_patterns: list[Pattern[str] | str] | None = None
    locale: Literal["en", "vi", "auto"] = "auto"

    @property
    def preferredLength(self) -> int | None:
        return self.preferred_length

    @property
    def allowAlphanumeric(self) -> bool:
        return self.allow_alphanumeric

    @property
    def customPatterns(self) -> list[Pattern[str] | str] | None:
        return self.custom_patterns


@dataclass(slots=True)
class OtpResult:
    """Extracted OTP result with confidence metrics and context."""

    code: str
    digits: int
    confidence: float
    pattern_matched: str
    raw_text: str
    context_snippet: str | None = None

    @property
    def patternMatched(self) -> str:
        return self.pattern_matched

    @property
    def contextSnippet(self) -> str | None:
        return self.context_snippet

    @property
    def rawText(self) -> str:
        return self.raw_text
