"""OTP extraction package export."""

from .patterns import (
    KEYWORDS_EN,
    KEYWORDS_VI,
    CONTEXT_PREFIX_REGEX,
    CONTEXT_SUFFIX_REGEX,
    NUMERIC_CODE_REGEXES,
    ALPHANUMERIC_CODE_REGEX,
    COPYRIGHT_YEAR_REGEX,
)
from .otp_extractor import (
    clean_html,
    extract_otp,
    extract_all_otps,
    cleanHtml,
    extractOtp,
    extractAllOtps,
)

__all__ = [
    "KEYWORDS_EN",
    "KEYWORDS_VI",
    "CONTEXT_PREFIX_REGEX",
    "CONTEXT_SUFFIX_REGEX",
    "NUMERIC_CODE_REGEXES",
    "ALPHANUMERIC_CODE_REGEX",
    "COPYRIGHT_YEAR_REGEX",
    "clean_html",
    "extract_otp",
    "extract_all_otps",
    "cleanHtml",
    "extractOtp",
    "extractAllOtps",
]
