"""Unit tests for OData Query Builder."""

from __future__ import annotations
import unittest
from datetime import datetime, timezone
from urllib.parse import unquote_plus
from outlook_graph.http.query_builder import (
    build_odata_query,
    escape_odata_string,
    get_mail_endpoint,
    buildODataQuery,
    escapeODataString,
    getMailEndpoint,
)
from outlook_graph.types.mail import GetMessagesOptions


class TestQueryBuilder(unittest.TestCase):
    """Test suite for OData query construction and endpoint resolution."""

    def test_escape_single_quotes_for_odata(self) -> None:
        self.assertEqual(escape_odata_string("O'Reilly"), "O''Reilly")
        self.assertEqual(escapeODataString("O'Reilly"), "O''Reilly")

    def test_return_empty_string_when_no_options(self) -> None:
        self.assertEqual(build_odata_query(None), "")
        self.assertEqual(build_odata_query(), "")
        self.assertEqual(buildODataQuery(), "")

    def test_construct_endpoint_urls(self) -> None:
        self.assertEqual(
            get_mail_endpoint(),
            "https://graph.microsoft.com/v1.0/me/messages",
        )
        self.assertEqual(
            get_mail_endpoint("inbox"),
            "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages",
        )
        self.assertEqual(
            getMailEndpoint("junkemail"),
            "https://graph.microsoft.com/v1.0/me/mailFolders/junkemail/messages",
        )

    def test_format_filter_options(self) -> None:
        opts = GetMessagesOptions(
            unread_only=True,
            subject_contains="Verification",
            from_contains="security@service.com",
            top=5,
        )
        query = build_odata_query(opts)
        decoded_query = unquote_plus(query)

        self.assertIn("$filter=", decoded_query)
        self.assertIn("isRead eq false", decoded_query)
        self.assertIn("contains(subject, 'Verification')", decoded_query)
        self.assertIn(
            "contains(from/emailAddress/address, 'security@service.com')",
            decoded_query,
        )
        self.assertIn("$top=5", decoded_query)
        self.assertIn("$orderby=receivedDateTime desc", decoded_query)

    def test_support_date_filter_and_custom_filter_override(self) -> None:
        dt = datetime(2026, 8, 25, 12, 0, 0, tzinfo=timezone.utc)
        opts = GetMessagesOptions(received_after=dt)
        query = build_odata_query(opts)
        decoded = unquote_plus(query)
        self.assertIn("receivedDateTime ge 2026-08-25T12:00:00Z", decoded)

        custom_opts = GetMessagesOptions(
            filter="hasAttachments eq true and importance eq 'high'"
        )
        custom_query = build_odata_query(custom_opts)
        decoded_custom = unquote_plus(custom_query)
        self.assertIn(
            "$filter=hasAttachments eq true and importance eq 'high'",
            decoded_custom,
        )


if __name__ == "__main__":
    unittest.main()
