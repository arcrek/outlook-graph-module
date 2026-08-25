"""Account parser for pipe-delimited credentials strings and files."""

from __future__ import annotations
from pathlib import Path
from ..types.account import AccountCredentials, ParseAccountOptions
from ..types.errors import AccountParseError


def parse_account_line(
    line: str,
    options: ParseAccountOptions | None = None,
    line_number: int | None = None,
) -> AccountCredentials:
    """Parse a single pipe-delimited account credential line.

    Expected format: email|password|refresh_token|client_id[|authority]
    """
    should_trim = options.trim if options is not None else True
    raw = line.strip() if should_trim else line

    if not raw:
        raise AccountParseError(
            "Cannot parse empty account line",
            raw_line=line,
            line_number=line_number,
        )

    parts = raw.split("|")
    if len(parts) < 4:
        raise AccountParseError(
            "Invalid account format. Expected 'email|password|refresh_token|client_id[|authority]'",
            raw_line=line,
            line_number=line_number,
        )

    email = (parts[0].strip() if should_trim else parts[0]) or ""
    password = parts[1].strip() if should_trim else parts[1]
    refresh_token = (parts[2].strip() if should_trim else parts[2]) or ""
    client_id = (parts[3].strip() if should_trim else parts[3]) or ""

    authority: str | None = None
    if len(parts) > 4 and parts[4]:
        authority = parts[4].strip() if should_trim else parts[4]
    elif options is not None and options.default_authority:
        authority = options.default_authority

    if not email:
        raise AccountParseError(
            "Account email is missing or empty",
            raw_line=line,
            line_number=line_number,
        )
    if not refresh_token:
        raise AccountParseError(
            "Account refresh token is missing or empty",
            raw_line=line,
            line_number=line_number,
        )
    if not client_id:
        raise AccountParseError(
            "Account client ID is missing or empty",
            raw_line=line,
            line_number=line_number,
        )

    return AccountCredentials(
        email=email,
        password=password,
        refresh_token=refresh_token,
        client_id=client_id,
        authority=authority,
    )


def parse_account_lines(
    content: str,
    options: ParseAccountOptions | None = None,
) -> list[AccountCredentials]:
    """Parse multi-line string content into a list of AccountCredentials."""
    ignore_empty = options.ignore_empty_lines if options is not None else True
    ignore_comments = options.ignore_comments if options is not None else True

    lines = content.splitlines()
    accounts: list[AccountCredentials] = []

    for idx, raw_line in enumerate(lines, start=1):
        line = raw_line.strip()
        if ignore_empty and not line:
            continue
        if ignore_comments and line.startswith("#"):
            continue

        accounts.append(
            parse_account_line(raw_line, options=options, line_number=idx)
        )

    return accounts


def parse_account_file(
    file_path: str | Path,
    options: ParseAccountOptions | None = None,
) -> list[AccountCredentials]:
    """Read a file from disk and parse all accounts inside it."""
    path = Path(file_path)
    if not path.is_file():
        raise FileNotFoundError(f"Account file not found: {file_path}")

    content = path.read_text(encoding="utf-8")
    return parse_account_lines(content, options=options)


# TypeScript naming parity aliases
parseAccountLine = parse_account_line
parseAccountLines = parse_account_lines
parseAccountFile = parse_account_file
