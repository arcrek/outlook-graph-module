"""OData v4 query string builder and endpoint resolver."""

from __future__ import annotations
from datetime import datetime
from urllib.parse import quote_plus
from ..types.mail import GetMessagesOptions


def escape_odata_string(value: str) -> str:
    """Escape single quotes for OData query literals."""
    return value.replace("'", "''")


def build_odata_query(options: GetMessagesOptions | None = None) -> str:
    """Build an encoded OData v4 query string from GetMessagesOptions."""
    if options is None:
        return ""

    params: list[tuple[str, str]] = []
    filter_clauses: list[str] = []

    if options.filter:
        filter_clauses.append(options.filter)
    else:
        if options.unread_only is True:
            filter_clauses.append("isRead eq false")

        if options.subject_contains:
            escaped = escape_odata_string(options.subject_contains)
            filter_clauses.append(f"contains(subject, '{escaped}')")

        if options.from_contains:
            escaped = escape_odata_string(options.from_contains)
            filter_clauses.append(
                f"contains(from/emailAddress/address, '{escaped}')"
            )

        if options.received_after:
            if isinstance(options.received_after, datetime):
                # Ensure UTC ISO format with Z
                dt = options.received_after
                if dt.tzinfo is not None:
                    iso_date = dt.isoformat().replace("+00:00", "Z")
                else:
                    iso_date = dt.strftime("%Y-%m-%dT%H:%M:%S.000Z")
            else:
                iso_date = str(options.received_after)
            filter_clauses.append(f"receivedDateTime ge {iso_date}")

    if filter_clauses:
        params.append(("$filter", " and ".join(filter_clauses)))

    if options.top is not None:
        clamped_top = max(1, min(1000, options.top))
        params.append(("$top", str(clamped_top)))

    if options.skip is not None:
        clamped_skip = max(0, options.skip)
        params.append(("$skip", str(clamped_skip)))

    if options.select and len(options.select) > 0:
        params.append(("$select", ",".join(options.select)))

    if options.order_by:
        params.append(("$orderby", options.order_by))
    elif not options.search:
        params.append(("$orderby", "receivedDateTime desc"))

    if options.search:
        escaped_search = options.search.replace('"', '\\"')
        params.append(("$search", f'"{escaped_search}"'))

    if not params:
        return ""

    encoded_pairs = [
        f"{quote_plus(k)}={quote_plus(v)}" for k, v in params
    ]
    return f"?{'&'.join(encoded_pairs)}"


def get_mail_endpoint(folder: str | None = None) -> str:
    """Return the Graph API endpoint for messages, optionally scoped to a folder."""
    if not folder:
        return "https://graph.microsoft.com/v1.0/me/messages"

    clean_folder = folder.strip().lower()
    return f"https://graph.microsoft.com/v1.0/me/mailFolders/{clean_folder}/messages"


# TypeScript naming parity aliases
escapeODataString = escape_odata_string
buildODataQuery = build_odata_query
getMailEndpoint = get_mail_endpoint
