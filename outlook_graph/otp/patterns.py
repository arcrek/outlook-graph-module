"""Regex patterns and keyword dictionaries for OTP extraction."""

from __future__ import annotations
import re
from dataclasses import dataclass
from typing import Pattern

KEYWORDS_EN: list[str] = [
    "verification code",
    "security code",
    "confirmation code",
    "one-time password",
    "one-time passcode",
    "otp",
    "passcode",
    "secret code",
    "access code",
    "login code",
    "authorization code",
    "auth code",
    "validation code",
    "pin code",
    "temporary password",
]

KEYWORDS_VI: list[str] = [
    "mã xác thực",
    "mã xác nhận",
    "mã otp",
    "mã bảo mật",
    "mã kiểm tra",
    "mã xác minh",
    "mật khẩu một lần",
    "mã kích hoạt",
    "mã an toàn",
    "mã giao dịch",
    "mã đăng nhập",
    "mã số bí mật",
]

_ALL_KEYWORDS = sorted(KEYWORDS_EN + KEYWORDS_VI, key=len, reverse=True)
_KEYWORD_OR_PATTERN = "|".join(re.escape(k) for k in _ALL_KEYWORDS)

CONTEXT_PREFIX_REGEX: Pattern[str] = re.compile(
    rf"(?:(?:{_KEYWORD_OR_PATTERN})[^0-9a-zA-Z]{{0,40}}?(?:is|là|:|=|-|\s)?\s*)([0-9]{{4,8}}|[0-9]{{3}}[-\s][0-9]{{3}}|[A-Z0-9]{{6,8}})\b",
    re.IGNORECASE,
)

CONTEXT_SUFFIX_REGEX: Pattern[str] = re.compile(
    r"\b([0-9]{4,8}|[0-9]{3}[-\s][0-9]{3}|[A-Z0-9]{6,8})\s*(?:is\s*(?:your|the)?\s*(?:verification|security|confirmation|otp)?\s*code|là\s*mã\s*xác\s*(?:thực|nhận|minh))\b",
    re.IGNORECASE,
)


@dataclass(slots=True)
class NumericPatternDef:
    pattern: Pattern[str]
    digits: int
    weight: float


NUMERIC_CODE_REGEXES: list[NumericPatternDef] = [
    NumericPatternDef(re.compile(r"\b([0-9]{6})\b"), 6, 0.9),
    NumericPatternDef(re.compile(r"\b([0-9]{3}[-\s][0-9]{3})\b"), 6, 0.95),
    NumericPatternDef(re.compile(r"\b([0-9]{4})\b"), 4, 0.7),
    NumericPatternDef(re.compile(r"\b([0-9]{8})\b"), 8, 0.8),
    NumericPatternDef(re.compile(r"\b([0-9]{5})\b"), 5, 0.65),
    NumericPatternDef(re.compile(r"\b([0-9]{7})\b"), 7, 0.65),
]

ALPHANUMERIC_CODE_REGEX: Pattern[str] = re.compile(r"\b([A-Z0-9]{6,8})\b")
FALSE_POSITIVE_YEAR_REGEX: Pattern[str] = re.compile(r"\b(19\d\d|20\d\d)\b")
COPYRIGHT_YEAR_REGEX: Pattern[str] = re.compile(
    r"(?:copyright|©|\(c\))\s*(?:19\d\d|20\d\d)", re.IGNORECASE
)
