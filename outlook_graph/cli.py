"""CLI tool for Microsoft Outlook Graph API module."""

from __future__ import annotations
import argparse
import json
import sys
from dataclasses import asdict, is_dataclass
from pathlib import Path
from typing import Any

from .auth.token_manager import OutlookTokenManager
from .mail.mail_client import OutlookMailClient
from .otp.otp_extractor import clean_html
from .parser.account_parser import parse_account_file, parse_account_line
from .types.account import AccountCredentials
from .types.mail import GetMessagesOptions, WaitForEmailOptions


def _serialize(obj: Any) -> Any:
    """Helper to serialize dataclasses and custom objects to JSON-compatible dicts."""
    if is_dataclass(obj) and not isinstance(obj, type):
        data = asdict(obj)
        # Convert bytes to string representation if needed
        if "data" in data and data["data"] is not None:
            data["data"] = f"<bytes: {len(data['data'])}>"
        return data
    if isinstance(obj, bytes):
        return f"<bytes: {len(obj)}>"
    if hasattr(obj, "__dict__"):
        return obj.__dict__
    return str(obj)


def resolve_account(args: argparse.Namespace) -> AccountCredentials:
    """Resolve account credentials from either command-line string or account file."""
    if args.account:
        return parse_account_line(args.account)

    file_path = args.file or "authenticated-test-account"
    accounts = parse_account_file(file_path)
    if not accounts:
        raise ValueError(f"No valid accounts found in file '{file_path}'")
    return accounts[0]


def build_parser() -> argparse.ArgumentParser:
    """Construct command-line argument parser."""
    parser = argparse.ArgumentParser(
        prog="outlook-graph",
        description="Microsoft Outlook Graph API CLI - Headless account automation & OTP extractor",
    )

    subparsers = parser.add_subparsers(dest="command", help="Command to execute")

    # Common arguments helper
    def add_common_args(subp: argparse.ArgumentParser) -> None:
        subp.add_argument(
            "-a",
            "--account",
            help="Pipe-delimited account string (email|password|refresh_token|client_id[|authority])",
        )
        subp.add_argument(
            "-f",
            "--file",
            help="Path to account file (default: authenticated-test-account)",
        )
        subp.add_argument(
            "--json", action="store_true", help="Output results in JSON format"
        )

    # 1. auth
    auth_p = subparsers.add_parser(
        "auth", help="Test OAuth token refresh for an account"
    )
    add_common_args(auth_p)

    # 2. list
    list_p = subparsers.add_parser(
        "list", help="List recent messages in mailbox"
    )
    add_common_args(list_p)
    list_p.add_argument(
        "-t", "--top", type=int, default=10, help="Max messages to fetch (default: 10)"
    )
    list_p.add_argument("--folder", help="Target folder (e.g. inbox, junkemail)")
    list_p.add_argument(
        "-u", "--unread", action="store_true", help="Filter unread messages only"
    )
    list_p.add_argument(
        "-s", "--subject", help="Filter messages containing subject"
    )
    list_p.add_argument(
        "--from", dest="from_filter", help="Filter messages from sender"
    )

    # 3. latest
    latest_p = subparsers.add_parser(
        "latest", help="Get the latest email matching criteria"
    )
    add_common_args(latest_p)
    latest_p.add_argument(
        "--folder", help="Target folder (e.g. inbox, junkemail)"
    )
    latest_p.add_argument(
        "-u", "--unread", action="store_true", help="Filter unread messages only"
    )
    latest_p.add_argument(
        "-s", "--subject", help="Filter messages containing subject"
    )
    latest_p.add_argument(
        "--from", dest="from_filter", help="Filter messages from sender"
    )

    # 4. read
    read_p = subparsers.add_parser("read", help="Read full email by message ID")
    add_common_args(read_p)
    read_p.add_argument(
        "--id", dest="message_id", required=True, help="Message ID to read"
    )
    read_p.add_argument(
        "--format",
        choices=["text", "html"],
        default="text",
        help="Body content format (default: text)",
    )

    # 5. wait
    wait_p = subparsers.add_parser(
        "wait", help="Poll and wait for incoming email matching criteria"
    )
    add_common_args(wait_p)
    wait_p.add_argument(
        "-s", "--subject", help="Filter incoming email by subject"
    )
    wait_p.add_argument(
        "--from", dest="from_filter", help="Filter incoming email by sender"
    )
    wait_p.add_argument(
        "--timeout",
        dest="timeout_sec",
        type=int,
        default=60,
        help="Timeout in seconds (default: 60)",
    )
    wait_p.add_argument(
        "-u", "--unread", action="store_true", help="Filter unread messages only"
    )

    # 6. otp
    otp_p = subparsers.add_parser(
        "otp", help="Extract OTP / verification code from email"
    )
    add_common_args(otp_p)
    otp_p.add_argument("--id", dest="message_id", help="Specific message ID")
    otp_p.add_argument(
        "-w",
        "--wait",
        action="store_true",
        help="Wait for incoming email before extracting OTP",
    )
    otp_p.add_argument(
        "-s", "--subject", help="Subject filter for matching email"
    )
    otp_p.add_argument(
        "--from", dest="from_filter", help="Sender filter for matching email"
    )
    otp_p.add_argument(
        "--timeout",
        dest="timeout_sec",
        type=int,
        default=60,
        help="Wait timeout in seconds (default: 60)",
    )

    # 7. test-all
    test_all_p = subparsers.add_parser(
        "test-all", help="Batch test all accounts in file"
    )
    test_all_p.add_argument(
        "-f",
        "--file",
        default="authenticated-test-account",
        help="Path to accounts file",
    )
    test_all_p.add_argument(
        "--json", action="store_true", help="Output results in JSON format"
    )

    return parser


def main() -> None:
    parser = build_parser()
    if len(sys.argv) < 2:
        parser.print_help()
        sys.exit(0)

    args = parser.parse_args()
    if not args.command:
        parser.print_help()
        sys.exit(0)

    client = OutlookMailClient()

    try:
        if args.command == "auth":
            account = resolve_account(args)
            print(f"[AUTH] Refreshing OAuth token for {account.email}...")
            token_resp = client.token_manager.refresh_access_token(account)
            if args.json:
                print(json.dumps(asdict(token_resp), indent=2))
            else:
                print("[SUCCESS] Access token obtained successfully!")
                print(f"  Token Type : {token_resp.token_type}")
                print(f"  Expires In : {token_resp.expires_in}s")
                print(f"  Scope      : {token_resp.scope}")
                preview = (
                    token_resp.access_token[:30] + "..."
                    if len(token_resp.access_token) > 30
                    else token_resp.access_token
                )
                print(f"  Access Token Preview: {preview}")

        elif args.command == "list":
            account = resolve_account(args)
            print(f"[LIST] Fetching messages for {account.email}...")
            opts = GetMessagesOptions(
                top=args.top,
                folder=args.folder,
                unread_only=args.unread,
                subject_contains=args.subject,
                from_contains=args.from_filter,
            )
            messages = client.get_messages(account, opts)

            if args.json:
                print(
                    json.dumps(
                        [_serialize(m) for m in messages], indent=2, default=_serialize
                    )
                )
            else:
                print(f"Found {len(messages)} message(s):\n")
                for idx, msg in enumerate(messages, start=1):
                    sender = (
                        f"{msg.from_recipient.email_address.name or ''} <{msg.from_recipient.email_address.address}>"
                        if msg.from_recipient and msg.from_recipient.email_address
                        else "(Unknown)"
                    )
                    read_status = "[READ]" if msg.is_read else "[UNREAD]"
                    print(f"{idx}. {read_status} {msg.subject}")
                    print(f"   From: {sender} | Date: {msg.received_date_time}")
                    print(f"   ID  : {msg.id}")
                    print(f"   Preview: {msg.body_preview or '(No preview)'}\n")

        elif args.command == "latest":
            account = resolve_account(args)
            opts = GetMessagesOptions(
                folder=args.folder,
                unread_only=args.unread,
                subject_contains=args.subject,
                from_contains=args.from_filter,
            )
            msg = client.get_latest_message(account, opts)

            if not msg:
                print("No messages found matching criteria.")
                return

            if args.json:
                print(json.dumps(_serialize(msg), indent=2, default=_serialize))
            else:
                sender = (
                    f"{msg.from_recipient.email_address.name or ''} <{msg.from_recipient.email_address.address}>"
                    if msg.from_recipient and msg.from_recipient.email_address
                    else "(Unknown)"
                )
                print("Latest Message:")
                print(f"Subject : {msg.subject}")
                print(f"From    : {sender}")
                print(f"Date    : {msg.received_date_time}")
                print(f"ID      : {msg.id}")
                print("\n--- Body Preview ---")
                print(msg.body_preview or clean_html(msg.body.content))

        elif args.command == "read":
            account = resolve_account(args)
            msg = client.get_message_by_id(account, args.message_id)
            attachments = (
                client.get_message_attachments(account, args.message_id)
                if msg.has_attachments
                else []
            )

            if args.json:
                payload = {
                    "message": _serialize(msg),
                    "attachments": [_serialize(a) for a in attachments],
                }
                print(json.dumps(payload, indent=2, default=_serialize))
            else:
                sender = (
                    f"{msg.from_recipient.email_address.name or ''} <{msg.from_recipient.email_address.address}>"
                    if msg.from_recipient and msg.from_recipient.email_address
                    else "(Unknown)"
                )
                print(f"Subject : {msg.subject}")
                print(f"From    : {sender}")
                print(f"Date    : {msg.received_date_time}")
                print(f"Attachments: {len(attachments)}")
                for att in attachments:
                    print(f"  - {att.name} ({att.size} bytes)")
                print(
                    f"\n--- Body Content ({args.format or msg.body.content_type}) ---"
                )
                print(
                    msg.body.content
                    if args.format == "html"
                    else clean_html(msg.body.content)
                )

        elif args.command == "wait":
            account = resolve_account(args)
            timeout_sec = args.timeout_sec or 60
            print(
                f"[WAIT] Polling for incoming email on {account.email} (timeout: {timeout_sec}s)..."
            )
            opts = WaitForEmailOptions(
                subject_contains=args.subject,
                from_contains=args.from_filter,
                timeout_ms=timeout_sec * 1000,
                unread_only=args.unread,
            )
            msg = client.wait_for_email(account, opts)

            if args.json:
                print(json.dumps(_serialize(msg), indent=2, default=_serialize))
            else:
                sender = (
                    msg.from_recipient.email_address.address
                    if msg.from_recipient and msg.from_recipient.email_address
                    else "Unknown"
                )
                print("\n[MATCHED] New email arrived!")
                print(f"Subject : {msg.subject}")
                print(f"From    : {sender}")
                print(f"Date    : {msg.received_date_time}")
                print(f"Preview : {msg.body_preview}")

        elif args.command == "otp":
            account = resolve_account(args)
            target_msg = None

            if args.message_id:
                target_msg = client.get_message_by_id(account, args.message_id)
            elif args.wait:
                timeout_sec = args.timeout_sec or 60
                print(
                    f"[OTP] Waiting for verification email on {account.email} (timeout: {timeout_sec}s)..."
                )
                target_msg = client.wait_for_email(
                    account,
                    WaitForEmailOptions(
                        subject_contains=args.subject,
                        from_contains=args.from_filter,
                        timeout_ms=timeout_sec * 1000,
                    ),
                )
            else:
                print(f"[OTP] Fetching latest email on {account.email}...")
                target_msg = client.get_latest_message(
                    account,
                    GetMessagesOptions(
                        subject_contains=args.subject,
                        from_contains=args.from_filter,
                    ),
                )

            if not target_msg:
                print("No email found to extract OTP from.")
                return

            otp = client.extract_otp(target_msg)
            if args.json:
                payload = {
                    "message": _serialize(target_msg),
                    "otp": _serialize(otp) if otp else None,
                }
                print(json.dumps(payload, indent=2, default=_serialize))
            else:
                print(f"\nEmail Subject : {target_msg.subject}")
                print(f"Email Date    : {target_msg.received_date_time}")
                if otp:
                    print("\n[OTP DETECTED]")
                    print(f"  Code       : {otp.code}")
                    print(f"  Confidence : {int(otp.confidence * 100)}%")
                    print(f"  Pattern    : {otp.pattern_matched}")
                    print(f'  Context    : "{otp.context_snippet}"')
                else:
                    print("\n[OTP] No verification code detected in this email.")

        elif args.command == "test-all":
            file_path = args.file or "authenticated-test-account"
            accounts = parse_account_file(file_path)
            print(
                f"\n=== Batch Testing {len(accounts)} Accounts from '{file_path}' ===\n"
            )

            results: list[dict[str, Any]] = []
            for i, acc in enumerate(accounts, start=1):
                sys.stdout.write(f"[{i}/{len(accounts)}] Testing {acc.email}... ")
                sys.stdout.flush()
                try:
                    token_resp = client.token_manager.refresh_access_token(acc)
                    messages = client.get_messages(acc, GetMessagesOptions(top=3))
                    sys.stdout.write(
                        f"OK (Token valid, {len(messages)} message(s) in inbox)\n"
                    )
                    results.append(
                        {
                            "email": acc.email,
                            "status": "success",
                            "message_count": len(messages),
                            "token_expires_in": token_resp.expires_in,
                        }
                    )
                except Exception as err:
                    sys.stdout.write(f"FAILED: {err}\n")
                    results.append(
                        {
                            "email": acc.email,
                            "status": "failed",
                            "error": str(err),
                        }
                    )

            passed_count = sum(1 for r in results if r["status"] == "success")
            print(f"\nSummary: {passed_count}/{len(accounts)} passed.")

    except Exception as err:
        print(f"\n[ERROR] {err}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
