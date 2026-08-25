"""Multi-language (EN/VI) OTP extraction and HTML cleaning engine."""

from __future__ import annotations
import html
import re
from typing import Any

from ..types.mail import OutlookMessage
from ..types.otp import OtpExtractOptions, OtpResult
from .patterns import (
    CONTEXT_PREFIX_REGEX,
    CONTEXT_SUFFIX_REGEX,
    COPYRIGHT_YEAR_REGEX,
    KEYWORDS_EN,
    KEYWORDS_VI,
    NUMERIC_CODE_REGEXES,
)


def clean_html(raw_html: str) -> str:
    """Clean HTML tags and decode HTML entities to plain text."""
    if not raw_html:
        return ""

    text = re.sub(r"<style[\s\S]*?</style>", " ", raw_html, flags=re.IGNORECASE)
    text = re.sub(r"<script[\s\S]*?</script>", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"<svg[\s\S]*?</svg>", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", " ", text)

    # Decode common HTML entities
    entities: dict[str, str] = {
        "&nbsp;": " ",
        "&amp;": "&",
        "&lt;": "<",
        "&gt;": ">",
        "&quot;": '"',
        "&#39;": "'",
        "&apos;": "'",
        "&zwnj;": "",
        "&zwj;": "",
    }
    for entity, replacement in entities.items():
        text = text.replace(entity, replacement)

    # Handle numeric decimal & hex entities
    def _replace_decimal(match: re.Match[str]) -> str:
        try:
            return chr(int(match.group(1), 10))
        except (ValueError, OverflowError):
            return match.group(0)

    def _replace_hex(match: re.Match[str]) -> str:
        try:
            return chr(int(match.group(1), 16))
        except (ValueError, OverflowError):
            return match.group(0)

    text = re.sub(r"&#(\d+);", _replace_decimal, text)
    text = re.sub(r"&#x([0-9a-fA-F]+);", _replace_hex, text)

    # General html unescape for any remaining entities
    text = html.unescape(text)

    # Collapse multiple whitespace to single space
    return re.sub(r"\s+", " ", text).strip()


def _resolve_raw_text(input_data: OutlookMessage | str) -> str:
    """Extract plain text string from either an OutlookMessage or a raw string."""
    if isinstance(input_data, str):
        if "<" in input_data and ">" in input_data:
            return clean_html(input_data)
        return input_data

    subject = input_data.subject or ""
    body_content = ""
    if input_data.body and input_data.body.content:
        body_content = input_data.body.content
    elif input_data.body_preview:
        body_content = input_data.body_preview

    if (
        input_data.body
        and input_data.body.content_type == "html"
        or ("<" in body_content and ">" in body_content)
    ):
        cleaned_body = clean_html(body_content)
    else:
        cleaned_body = body_content

    return f"{subject}\n{cleaned_body}"


def extract_otp(
    input_data: OutlookMessage | str,
    options: OtpExtractOptions | None = None,
) -> OtpResult | None:
    """Extract single highest-confidence OTP from message or text."""
    results = extract_all_otps(input_data, options)
    return results[0] if results else None


def extract_all_otps(
    input_data: OutlookMessage | str,
    options: OtpExtractOptions | None = None,
) -> list[OtpResult]:
    """Extract all candidate OTPs scored and ranked by confidence."""
    raw_text = _resolve_raw_text(input_data)
    if not raw_text.strip():
        return []

    candidates: list[OtpResult] = []
    seen_codes: set[str] = set()

    # 1. Direct context prefix regex
    prefix_match = CONTEXT_PREFIX_REGEX.search(raw_text)
    if prefix_match and prefix_match.group(1):
        raw_code = prefix_match.group(1)
        normalized_code = re.sub(r"[-\s]", "", raw_code)
        index = prefix_match.start()
        snippet_start = max(0, index - 30)
        snippet_end = min(len(raw_text), index + len(prefix_match.group(0)) + 30)
        snippet = raw_text[snippet_start:snippet_end].strip()

        seen_codes.add(normalized_code)
        candidates.append(
            OtpResult(
                code=normalized_code,
                digits=len(normalized_code),
                confidence=0.98,
                pattern_matched="context_prefix",
                context_snippet=snippet,
                raw_text=raw_code,
            )
        )

    # 2. Direct context suffix regex
    suffix_match = CONTEXT_SUFFIX_REGEX.search(raw_text)
    if suffix_match and suffix_match.group(1):
        raw_code = suffix_match.group(1)
        normalized_code = re.sub(r"[-\s]", "", raw_code)
        if normalized_code not in seen_codes:
            index = suffix_match.start()
            snippet_start = max(0, index - 30)
            snippet_end = min(
                len(raw_text), index + len(suffix_match.group(0)) + 30
            )
            snippet = raw_text[snippet_start:snippet_end].strip()

            seen_codes.add(normalized_code)
            candidates.append(
                OtpResult(
                    code=normalized_code,
                    digits=len(normalized_code),
                    confidence=0.95,
                    pattern_matched="context_suffix",
                    context_snippet=snippet,
                    raw_text=raw_code,
                )
            )

    # 3. Keyword proximity search with general numeric patterns
    locale = options.locale if options else "auto"
    if locale == "en":
        all_keywords = KEYWORDS_EN
    elif locale == "vi":
        all_keywords = KEYWORDS_VI
    else:
        all_keywords = KEYWORDS_EN + KEYWORDS_VI

    lower_text = raw_text.lower()

    for item in NUMERIC_CODE_REGEXES:
        pattern = item.pattern
        digits = item.digits
        weight = item.weight

        if options and options.preferred_length and options.preferred_length != digits:
            continue

        for match in pattern.finditer(raw_text):
            raw_code = match.group(1)
            normalized_code = re.sub(r"[-\s]", "", raw_code)
            match_index = match.start()

            # Filter out years (1900-2099) if part of copyright notice
            is_year = digits == 4 and (
                normalized_code.startswith("19") or normalized_code.startswith("20")
            )
            if is_year:
                surrounding_text = raw_text[
                    max(0, match_index - 40) : min(len(raw_text), match_index + 40)
                ]
                if COPYRIGHT_YEAR_REGEX.search(surrounding_text):
                    continue

            # Calculate proximity distance to closest keyword
            min_distance = float("inf")
            closest_keyword = ""

            for kw in all_keywords:
                kw_pos = lower_text.find(kw)
                while kw_pos != -1:
                    dist = abs(match_index - kw_pos)
                    if dist < min_distance:
                        min_distance = dist
                        closest_keyword = kw
                    kw_pos = lower_text.find(kw, kw_pos + 1)

            confidence = weight
            if min_distance < 60:
                confidence = min(0.99, weight + 0.15)
            elif min_distance < 150:
                confidence = min(0.85, weight + 0.05)
            elif min_distance < 300:
                confidence = weight * 0.7
            else:
                confidence = 0.1 if is_year else weight * 0.4

            if normalized_code in seen_codes:
                # Upgrade existing candidate if higher confidence
                for c in candidates:
                    if c.code == normalized_code and confidence > c.confidence:
                        c.confidence = round(confidence, 2)
                continue

            seen_codes.add(normalized_code)
            snippet_start = max(0, match_index - 30)
            snippet_end = min(len(raw_text), match_index + len(raw_code) + 30)
            snippet = raw_text[snippet_start:snippet_end].strip()

            candidates.append(
                OtpResult(
                    code=normalized_code,
                    digits=len(normalized_code),
                    confidence=round(confidence, 2),
                    pattern_matched=f"numeric_{digits}_{'proximate' if closest_keyword else 'isolated'}",
                    context_snippet=snippet,
                    raw_text=raw_code,
                )
            )

    # 4. Custom regex patterns if provided
    if options and options.custom_patterns:
        for custom_reg in options.custom_patterns:
            compiled = (
                custom_reg
                if hasattr(custom_reg, "finditer")
                else re.compile(str(custom_reg))
            )
            for match in compiled.finditer(raw_text):
                code = match.group(1) if match.groups() else match.group(0)
                if code not in seen_codes:
                    seen_codes.add(code)
                    candidates.append(
                        OtpResult(
                            code=code,
                            digits=len(code),
                            confidence=0.9,
                            pattern_matched="custom_pattern",
                            context_snippet=match.group(0),
                            raw_text=match.group(0),
                        )
                    )

    # Sort candidates by confidence descending, then digits descending
    candidates.sort(key=lambda c: (-c.confidence, -c.digits))
    return candidates


# TypeScript naming parity aliases
cleanHtml = clean_html
extractOtp = extract_otp
extractAllOtps = extract_all_otps
