# Microsoft Outlook Graph API - Python Guide

A high-performance, enterprise-grade Python module for interacting with Microsoft Outlook and Hotmail mailboxes via the Microsoft Graph API v1.0. Built for account automation, inbox monitoring, verification workflows, and intelligent multi-language OTP extraction using pipe-delimited account credentials.

[![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)](https://www.python.org/)
[![Zero External Runtime Dependencies](https://img.shields.io/badge/Dependencies-Zero%20Runtime%20Deps-orange.svg)]()
[![License](https://img.shields.io/badge/License-MIT-purple.svg)](../LICENSE)

---

## Table of Contents

- [Key Features](#key-features)
- [Zero External Runtime Dependency Architecture](#zero-external-runtime-dependency-architecture)
- [Installation & Quickstart](#installation--quickstart)
  - [Install via pip](#install-via-pip)
  - [Direct Module Import](#direct-module-import)
- [Account Format Specification](#account-format-specification)
- [Python API Usage & Code Examples](#python-api-usage--code-examples)
  - [1. Parsing Account Credentials](#1-parsing-account-credentials)
  - [2. OAuth2 Token Management & Rotation](#2-oauth2-token-management--rotation)
  - [3. Fetching & Filtering Messages](#3-fetching--filtering-messages)
  - [4. Reading Email Body & Attachments](#4-reading-email-body--attachments)
  - [5. Polling with `wait_for_email`](#5-polling-with-wait_for_email)
  - [6. Multi-Language OTP Extraction (EN & VI)](#6-multi-language-otp-extraction-en--vi)
  - [7. Modifying Messages (Mark as Read / Delete)](#7-modifying-messages-mark-as-read--delete)
- [Command Line Interface (CLI) Guide](#command-line-interface-cli-guide)
  - [Global Options & Flags](#global-options--flags)
  - [Command Reference](#command-reference)
    - [`auth`](#auth---test-oauth2-token-refresh)
    - [`list`](#list---list-messages-from-inbox-or-folder)
    - [`latest`](#latest---get-latest-email)
    - [`read`](#read---read-email-content--attachments)
    - [`wait`](#wait---poll-for-matching-incoming-email)
    - [`otp`](#otp---extract-verification-code)
    - [`test-all`](#test-all---batch-account-validation)
- [Unit & Integration Testing Guide](#unit--integration-testing-guide)
- [Error Handling & Exceptions](#error-handling--exceptions)
- [Architecture & Design Principles](#architecture--design-principles)

---

## Key Features

- **Zero External Runtime Dependencies**: Built entirely using Python standard library packages (`urllib.request`, `json`, `dataclasses`, `re`, `html`, `time`, `threading`, `random`, `base64`, `argparse`, `pathlib`). No third-party dependencies required.
- **Automated OAuth 2.0 Token Lifecycle**:
  - Headless token refresh using OAuth 2.0 `refresh_token` grant.
  - In-memory caching with configurable safety buffer (default 300s).
  - Thread-safe in-flight request de-duplication to prevent token stampedes across concurrent threads.
  - Rolling token rotation callbacks (`on_token_rotated`).
- **Resilience & Rate Limit Handling**:
  - Transparent HTTP 429 (`Retry-After`) exponential backoff with randomized jitter.
  - Automatic HTTP 401 re-authentication retry.
  - Transient HTTP 5xx server error recovery.
- **Rich Mailbox Operations**:
  - OData v4 query support: `$filter`, `$select`, `$top`, `$skip`, `$orderby`, and `$search`.
  - Folder scoping (Inbox, Junk Email, custom folders).
  - HTML entity decoding, sanitization, and plain-text conversion.
  - Attachment downloading and decoding to native `bytes`.
  - Status updates (`mark_as_read`, `delete_message`).
- **Non-Blocking Polling Engine (`wait_for_email`)**:
  - Long-polling loop with custom intervals, timeouts, and `received_after` thresholds.
  - Custom predicate functions for complex matching criteria.
  - Optional automatic mark-as-read upon successful match.
- **Intelligent Multi-Language OTP Extractor**:
  - Multi-tier heuristic engine extracting 4–8 digit verification codes, hyphenated codes (`123-456`), and alphanumeric tokens.
  - Comprehensive keyword dictionaries in English and Vietnamese (`verification code`, `security code`, `mã xác thực`, `mã OTP`, etc.).
  - False-positive filtering for copyright years (1900–2099), dates, and order IDs.
  - Confidence scoring (0.0 to 1.0) and context snippet extraction.

---

## Zero External Runtime Dependency Architecture

The module is engineered to run in any Python 3.10+ environment without installing any packages via `pip`.

| Standard Library Module | Responsibility |
| :--- | :--- |
| `urllib.request`, `urllib.parse`, `urllib.error` | HTTP networking, query string serialization, and error handling |
| `dataclasses` | Strongly typed domain entities (`@dataclass(slots=True)`) |
| `threading` | Thread-safe locks and event synchronization for in-flight token refresh de-duplication |
| `html`, `re` | HTML entity decoding, tag sanitization, and multilingual OTP regex matching |
| `base64` | Email attachment binary decoding to Python `bytes` |
| `argparse`, `sys`, `json` | Command-line interface and structured JSON output |

---

## Installation & Quickstart

### Prerequisites
- **Python**: Version `3.10` or higher (`3.10`, `3.11`, `3.12`, `3.13`, `3.14`).

### Install via pip
Install locally in development or editable mode:

```bash
# Standard installation
pip install .

# Editable / development mode
pip install -e .
```

### Direct Module Import
Because the package has zero external runtime dependencies, you can also place the `outlook_graph` folder directly into your project:

```bash
# Direct execution without pip install
python3 -m outlook_graph --help
```

---

## Account Format Specification

The module uses a standardized pipe-delimited format for account credentials:

```text
email|password|refresh_token|client_id[|authority]
```

### Field Breakdown

| Field Index | Field Name | Required | Description | Example |
| :--- | :--- | :--- | :--- | :--- |
| `0` | `email` | **Yes** | Hotmail / Outlook / Microsoft 365 email address | `user@hotmail.com` |
| `1` | `password` | No | Account password (stored for convenience; not used in OAuth flow) | `P@ssw0rd123` |
| `2` | `refresh_token` | **Yes** | Valid OAuth 2.0 Refresh Token issued by Microsoft Identity | `M.C504_BL2.0.U.-C...` |
| `3` | `client_id` | **Yes** | Azure AD Application (Client) ID | `9e5f94bc-e8a4-4e73-b8be-63364c29d753` |
| `4` | `authority` | No | Tenant authority (defaults to `common` or `consumers`) | `common` |

### Account File Conventions
When reading accounts from a file (e.g. `authenticated-test-account`):
- **Empty lines** and whitespace-only lines are ignored.
- **Comment lines** starting with `#` are ignored.
- Leading and trailing whitespace around each field is trimmed.

```text
# Development Accounts
user1@hotmail.com|pwd1|M.C504_BL2.0.U.token1|9e5f94bc-e8a4-4e73-b8be-63364c29d753
user2@outlook.com|pwd2|M.C504_BL2.0.U.token2|9e5f94bc-e8a4-4e73-b8be-63364c29d753|consumers
```

---

## Python API Usage & Code Examples

### 1. Parsing Account Credentials

```python
from outlook_graph import (
    parse_account_line,
    parse_account_lines,
    parse_account_file,
    AccountCredentials,
)

# 1. Parse a single account string
raw_line = "user@hotmail.com|secret_pwd|M.C504_BL2.0...|9e5f94bc-e8a4-4e73-b8be-63364c29d753"
account: AccountCredentials = parse_account_line(raw_line)
print(f"Parsed: {account.email}, Client ID: {account.client_id}")

# 2. Parse multi-line string content
raw_accounts = """
# Production accounts
user1@hotmail.com|pwd1|token1|client_id_1
user2@outlook.com|pwd2|token2|client_id_2|consumers
"""
accounts = parse_account_lines(raw_accounts)
print(f"Loaded {len(accounts)} accounts from string.")

# 3. Parse accounts directly from file
file_accounts = parse_account_file("./authenticated-test-account")
print(f"Loaded {len(file_accounts)} accounts from file.")
```

### 2. OAuth2 Token Management & Rotation

```python
from outlook_graph import (
    OutlookTokenManager,
    TokenManagerOptions,
    AccountCredentials,
    parse_account_line,
)

account = parse_account_line("user@hotmail.com|pwd|refresh_token|client_id")

# Optional callback for rolling refresh token rotation
def handle_token_rotated(old_token: str, new_token: str, acc: AccountCredentials):
    print(f"Token rotated for {acc.email}! New token: {new_token[:15]}...")
    # Persist updated refresh token to database or credentials file

options = TokenManagerOptions(
    authority="common",
    expiry_buffer_sec=300,  # 5-minute safety buffer before token expiration
    max_retries=3,
    on_token_rotated=handle_token_rotated,
)

token_manager = OutlookTokenManager(options)

# Retrieve access token (uses cache or triggers HTTP refresh)
access_token = token_manager.get_access_token(account)
print(f"Access Token: {access_token[:20]}...")

# Force immediate token refresh
fresh_token = token_manager.get_access_token(account, force_refresh=True)
```

### 3. Fetching & Filtering Messages

```python
from outlook_graph import (
    OutlookMailClient,
    GetMessagesOptions,
    parse_account_line,
)

account = parse_account_line("user@hotmail.com|pwd|refresh_token|client_id")
client = OutlookMailClient()

# List recent 10 messages from Inbox
messages = client.get_messages(
    account,
    options=GetMessagesOptions(
        folder="inbox",
        top=10,
        unread_only=False,
    ),
)

for msg in messages:
    sender = msg.from_recipient.email_address.address if msg.from_recipient else "Unknown"
    print(f"[{msg.id}] {msg.subject} from {sender} at {msg.received_date_time}")

# Advanced filtering with OData parameters
filtered = client.get_messages(
    account,
    options=GetMessagesOptions(
        folder="inbox",
        unread_only=True,
        subject_contains="Security code",
        from_contains="microsoft.com",
        top=5,
        order_by="receivedDateTime desc",
        select=["id", "subject", "from", "receivedDateTime", "bodyPreview"],
    ),
)
```

### 4. Reading Email Body & Attachments

```python
from outlook_graph import OutlookMailClient, clean_html, parse_account_line

account = parse_account_line("user@hotmail.com|pwd|refresh_token|client_id")
client = OutlookMailClient()

# Fetch message by ID (includes complete body content)
message = client.get_message_by_id(account, "AAMkAGI2...")

print(f"Subject: {message.subject}")
print(f"Content-Type: {message.body.content_type}")

# Convert HTML body to clean plain text
plain_text = clean_html(message.body.content)
print(f"Body text:\n{plain_text}")

# Retrieve and decode attachments
if message.has_attachments:
    attachments = client.get_message_attachments(account, message.id)
    for att in attachments:
        print(f"Attachment: {att.name} ({att.size} bytes, type: {att.content_type})")
        if att.data:
            # att.data is decoded Python bytes
            print(f"Decoded data size: {len(att.data)} bytes")
```

### 5. Polling with `wait_for_email`

`wait_for_email` polls the mailbox until an incoming email arrives that matches all filter criteria and custom predicates.

```python
from datetime import datetime, timezone, timedelta
from outlook_graph import (
    OutlookMailClient,
    WaitForEmailOptions,
    parse_account_line,
)

account = parse_account_line("user@hotmail.com|pwd|refresh_token|client_id")
client = OutlookMailClient()

print("Waiting for incoming verification email...")

incoming_email = client.wait_for_email(
    account,
    options=WaitForEmailOptions(
        subject_contains="Verification Code",
        from_contains="auth@service.com",
        timeout_ms=60000,   # 60s timeout
        interval_ms=3000,   # Check every 3s
        received_after=datetime.now(timezone.utc) - timedelta(seconds=30),
        mark_as_read_after_match=True, # Automatically mark as read
        predicate=lambda msg: "urgent" not in msg.subject.lower(),
    ),
)

print(f"Received email: {incoming_email.subject}")
print(f"Sender: {incoming_email.from_recipient.email_address.address}")
```

### 6. Multi-Language OTP Extraction (EN & VI)

The intelligent OTP extractor parses English and Vietnamese emails, handling HTML sanitization, numeric/hyphenated codes, and confidence ranking.

```python
from outlook_graph import (
    OutlookMailClient,
    extract_otp,
    extract_all_otps,
    OtpExtractOptions,
    parse_account_line,
)

account = parse_account_line("user@hotmail.com|pwd|refresh_token|client_id")
client = OutlookMailClient()

latest_email = client.get_latest_message(account)

if latest_email:
    # 1. Extract OTP directly from OutlookMessage
    otp_res = client.extract_otp(latest_email)
    if otp_res:
        print(f"Found OTP Code : {otp_res.code}")
        print(f"Digits         : {otp_res.digits}")
        print(f"Confidence     : {otp_res.confidence * 100:.1f}%")
        print(f"Pattern Matched: {otp_res.pattern_matched}")
        print(f"Context        : '{otp_res.context_snippet}'")

    # 2. Extract from raw string with options
    raw_vi_text = "Mã xác thực Shopee của bạn là 810492. Mã có hiệu lực trong 5 phút."
    vi_result = extract_otp(
        raw_vi_text,
        options=OtpExtractOptions(locale="vi", preferred_length=6),
    )
    if vi_result:
        print(f"Vietnamese OTP: {vi_result.code}")

    # 3. Extract all candidates
    all_candidates = extract_all_otps(
        "Code 1: 123-456, Security Code 2: 789012",
        options=OtpExtractOptions(allow_alphanumeric=False),
    )
    for cand in all_candidates:
        print(f"Candidate: {cand.code} (Confidence: {cand.confidence:.2f})")
```

### 7. Modifying Messages (Mark as Read / Delete)

```python
from outlook_graph import OutlookMailClient, parse_account_line

account = parse_account_line("user@hotmail.com|pwd|refresh_token|client_id")
client = OutlookMailClient()

msg_id = "AAMkAGI2..."

# Mark message as read / unread
client.mark_as_read(account, msg_id, is_read=True)
print("Marked message as read.")

# Delete message (moves to Deleted Items)
client.delete_message(account, msg_id)
print("Deleted message.")
```

---

## Command Line Interface (CLI) Guide

The CLI is executable via `python3 -m outlook_graph` or the `outlook-graph` command when installed.

```bash
python3 -m outlook_graph <command> [options]
```

### Global Options & Flags

| Flag | Description | Default |
| :--- | :--- | :--- |
| `-a`, `--account` | Account string: `email\|pwd\|refresh_token\|client_id[\|authority]` | `None` |
| `-f`, `--file` | Path to account credentials file | `authenticated-test-account` |
| `--json` | Enable formatted JSON output mode | `False` |
| `-h`, `--help` | Display command help and usage info | |

---

### Command Reference

#### `auth` - Test OAuth2 Token Refresh
Validates credentials and retrieves a fresh access token.

```bash
# Test account from file
python3 -m outlook_graph auth -f authenticated-test-account

# Test with raw account string and output JSON
python3 -m outlook_graph auth -a "user@hotmail.com|pwd|token|client_id" --json
```

#### `list` - List Messages from Inbox or Folder
Lists messages matching query filters.

```bash
# List top 5 messages
python3 -m outlook_graph list --top 5

# Filter by subject and unread status
python3 -m outlook_graph list -s "Security" --unread

# List messages in Junk Email folder
python3 -m outlook_graph list --folder junkemail --top 10 --json
```

#### `latest` - Get Latest Email
Fetches the single most recent email matching the filters.

```bash
# Get latest message
python3 -m outlook_graph latest

# Get latest message from specific sender
python3 -m outlook_graph latest --from "noreply@github.com" --json
```

#### `read` - Read Email Content & Attachments
Fetches full body and attachments for a specific message ID.

```bash
# Read plain-text email body
python3 -m outlook_graph read --id "AAMkAGI2..."

# Read raw HTML email body
python3 -m outlook_graph read --id "AAMkAGI2..." --format html

# Output full payload with attachments metadata in JSON
python3 -m outlook_graph read --id "AAMkAGI2..." --json
```

#### `wait` - Poll for Matching Incoming Email
Long-polls the mailbox until a matching email arrives.

```bash
# Wait for email with subject "Verification Code" within 60s
python3 -m outlook_graph wait -s "Verification Code" --timeout 60

# Wait for email from specific sender
python3 -m outlook_graph wait --from "auth@service.com" --timeout 120 --json
```

#### `otp` - Extract Verification Code
Extracts OTP verification code from the latest email, a specific message ID, or an incoming email.

```bash
# Extract OTP from latest email
python3 -m outlook_graph otp

# Wait for incoming verification email and extract OTP immediately
python3 -m outlook_graph otp -w -s "Verification" --timeout 60

# Extract OTP from specific message ID with JSON output
python3 -m outlook_graph otp --id "AAMkAGI2..." --json
```

#### `test-all` - Batch Account Validation
Iterates over all accounts in an accounts file and reports token refresh and mailbox listing status.

```bash
python3 -m outlook_graph test-all -f authenticated-test-account
```

---

## Unit & Integration Testing Guide

### Running Unit Tests

The test suite runs using Python's standard `unittest` framework with no extra dependencies:

```bash
# Run all unit tests
python3 -m unittest discover -s tests/unit

# Run specific unit test module
python3 -m unittest tests/unit/test_otp_extractor.py
python3 -m unittest tests/unit/test_token_manager.py
python3 -m unittest tests/unit/test_account_parser.py
python3 -m unittest tests/unit/test_graph_client.py
python3 -m unittest tests/unit/test_query_builder.py
python3 -m unittest tests/unit/test_mail_client.py
```

### Running Live Integration Tests

Integration tests validate against live Microsoft Graph and Microsoft Identity endpoints using credentials from `authenticated-test-account`:

```bash
# Run live integration tests
python3 -m unittest discover -s tests/integration
```

---

## Error Handling & Exceptions

All exceptions raised by the module inherit from `GraphModuleError`:

```
GraphModuleError (base exception)
├── AccountParseError     (Raised on invalid account strings/files)
├── TokenRefreshError      (Raised on OAuth2 token refresh failures)
├── GraphApiError          (Raised when Graph API returns HTTP errors)
├── RateLimitError         (Raised when HTTP 429 retries are exhausted)
└── TimeoutError           (Raised when wait_for_email exceeds timeout)
```

### Exception Handling Example

```python
from outlook_graph import (
    OutlookMailClient,
    GraphApiError,
    TokenRefreshError,
    RateLimitError,
    TimeoutError,
    parse_account_line,
)

try:
    account = parse_account_line("user@hotmail.com|pwd|token|client_id")
    client = OutlookMailClient()
    msg = client.get_latest_message(account)
except TokenRefreshError as e:
    print(f"Authentication failed: {e.message} (HTTP {e.status_code})")
except RateLimitError as e:
    print(f"Rate limited. Retry after {e.retry_after_ms}ms")
except GraphApiError as e:
    print(f"Graph API Error: {e.message} (Code: {e.graph_error_code})")
except TimeoutError as e:
    print(f"Email polling timed out: {e.message}")
```

---

## Architecture & Design Principles

1. **Thread-Safe In-Flight De-Duplication**: Prevents multiple threads from requesting redundant token refreshes for the same account simultaneously using `threading.Lock` and `threading.Event`.
2. **Exponential Backoff with Full Jitter**:
   $$t_{\text{wait}} = \begin{cases} (\text{Retry-After} \times 1000) + \text{uniform}(0, 200)\text{ms} & \text{if Header present} \\ (\text{initialDelay} \times 2^{\text{attempt}-1}) + \text{uniform}(0, 200)\text{ms} & \text{otherwise} \end{cases}$$
3. **Transparent 401 Re-Authentication**: Proactively acquires a fresh access token on HTTP 401 and replays the request once before propagating errors.
4. **4-Tier Intelligent OTP Heuristics**: Combines prefix context, suffix context, proximity weighting, and false-positive year/date suppression for accuracy in English and Vietnamese emails.
