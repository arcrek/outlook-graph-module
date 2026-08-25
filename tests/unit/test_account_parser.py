"""Unit tests for Account Parser."""

from __future__ import annotations
import unittest
from pathlib import Path
from outlook_graph.parser.account_parser import (
    parse_account_line,
    parse_account_lines,
    parse_account_file,
    parseAccountLine,
    parseAccountLines,
    parseAccountFile,
)
from outlook_graph.types.errors import AccountParseError


class TestAccountParser(unittest.TestCase):
    """Test suite for pipe-delimited account string & file parser."""

    def test_parse_valid_4_part_account_line(self) -> None:
        line = "user@hotmail.com|password123|M.C557_BAY.token...|9e5f94bc-e8a4-4e73-b8be-63364c29d753"
        acc = parse_account_line(line)

        self.assertEqual(acc.email, "user@hotmail.com")
        self.assertEqual(acc.password, "password123")
        self.assertEqual(acc.refresh_token, "M.C557_BAY.token...")
        self.assertEqual(acc.refreshToken, "M.C557_BAY.token...")
        self.assertEqual(acc.client_id, "9e5f94bc-e8a4-4e73-b8be-63364c29d753")
        self.assertEqual(acc.clientId, "9e5f94bc-e8a4-4e73-b8be-63364c29d753")
        self.assertIsNone(acc.authority)

    def test_parse_optional_authority_in_5_part_line(self) -> None:
        line = "user@hotmail.com|password123|token|client-id|consumers"
        acc = parse_account_line(line)
        self.assertEqual(acc.authority, "consumers")

    def test_throw_on_empty_or_whitespace_line(self) -> None:
        with self.assertRaises(AccountParseError):
            parse_account_line("")
        with self.assertRaises(AccountParseError):
            parse_account_line("   ")

    def test_throw_on_fewer_than_4_fields(self) -> None:
        with self.assertRaises(AccountParseError):
            parse_account_line("user@hotmail.com|pass")
        with self.assertRaises(AccountParseError):
            parse_account_line("user@hotmail.com|pass|only_three_parts")

    def test_throw_on_missing_required_fields(self) -> None:
        with self.assertRaisesRegex(AccountParseError, "email is missing"):
            parse_account_line("|pass|refresh_token|client_id")
        with self.assertRaisesRegex(AccountParseError, "refresh token is missing"):
            parse_account_line("user@hotmail.com|pass||client_id")
        with self.assertRaisesRegex(AccountParseError, "client ID is missing"):
            parse_account_line("user@hotmail.com|pass|refresh_token|")

    def test_parse_multi_line_ignoring_comments_and_blanks(self) -> None:
        content = """
# Hotmail Test Accounts
user1@hotmail.com|p1|rt1|cid1

# Secondary account
user2@hotmail.com|p2|rt2|cid2
"""
        accounts = parse_account_lines(content)
        self.assertEqual(len(accounts), 2)
        self.assertEqual(accounts[0].email, "user1@hotmail.com")
        self.assertEqual(accounts[1].email, "user2@hotmail.com")

    def test_parse_account_file_if_exists(self) -> None:
        file_path = Path("authenticated-test-account")
        if file_path.exists():
            accounts = parse_account_file(file_path)
            self.assertGreaterEqual(len(accounts), 1)
            self.assertTrue(accounts[0].email)
            self.assertTrue(accounts[0].client_id)

    def test_typescript_parity_aliases(self) -> None:
        line = "u@h.com|p|r|c"
        acc = parseAccountLine(line)
        self.assertEqual(acc.email, "u@h.com")
        accs = parseAccountLines(line)
        self.assertEqual(len(accs), 1)


if __name__ == "__main__":
    unittest.main()
