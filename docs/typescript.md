# Microsoft Outlook Graph API - TypeScript Guide

A high-performance, enterprise-grade TypeScript/Node.js module for interacting with Microsoft Outlook and Hotmail mailboxes via the Microsoft Graph API v1.0. Built for account automation, inbox monitoring, verification workflows, and intelligent multi-language OTP extraction using pipe-delimited account credentials.

[![TypeScript](https://img.shields.io/badge/TypeScript-5.7+-blue.svg)](https://www.typescriptlang.org/)
[![Zero External Runtime Dependencies](https://img.shields.io/badge/Dependencies-Zero%20Runtime%20Deps-orange.svg)]()
[![License](https://img.shields.io/badge/License-MIT-purple.svg)](../LICENSE)

---

## Table of Contents

- [Key Features](#key-features)
- [Installation & Setup](#installation--setup)
- [Account Format Specification](#account-format-specification)
- [Quick Start Guide](#quick-start-guide)
  - [1. Parsing Account Credentials](#1-parsing-account-credentials)
  - [2. Fetching & Filtering Messages](#2-fetching--filtering-messages)
  - [3. Reading Email Body & Attachments](#3-reading-email-body--attachments)
  - [4. Polling & Waiting for Incoming Emails](#4-polling--waiting-for-incoming-emails)
  - [5. Extracting OTP / Verification Codes](#5-extracting-otp--verification-codes)
  - [6. Managing Messages (Mark as Read / Delete)](#6-managing-messages-mark-as-read--delete)
- [CLI Usage Guide](#cli-usage-guide)
  - [CLI Options & Flags](#cli-options--flags)
  - [Command Reference & Examples](#command-reference--examples)
- [API Reference](#api-reference)
  - [`OutlookMailClient`](#outlookmailclient)
  - [`OutlookTokenManager`](#outlooktokenmanager)
  - [Account Parser Functions](#account-parser-functions)
  - [OTP Extractor Functions](#otp-extractor-functions)
  - [Query Builder Functions](#query-builder-functions)
  - [Error Hierarchy](#error-hierarchy)
- [TypeScript Types & Interfaces](#typescript-types--interfaces)
- [Testing](#testing)

---

## Key Features

- **Zero Heavy Runtime Dependencies**: Built entirely on native Node.js 18+ runtime APIs (`fetch`, `AbortSignal`, `Buffer`, `URLSearchParams`). No bloated SDKs like `@microsoft/microsoft-graph-client` or `@azure/msal-node`.
- **Automated OAuth 2.0 Token Management**:
  - Headless token refresh using OAuth 2.0 `refresh_token` grant.
  - Smart in-memory token caching with configurable expiration buffer (default 300s).
  - Rolling token rotation support with persistent callback hooks (`onTokenRotated`).
  - In-flight request de-duplication to prevent token stampedes.
- **Built-in Resilience & Rate Limiting**:
  - Transparent HTTP 429 (`Retry-After`) handling with exponential backoff and randomized jitter.
  - Automatic HTTP 401 re-authentication retry.
  - Transient HTTP 5xx server error recovery.
- **Rich Mailbox Operations**:
  - Full OData v4 query support: `$filter`, `$select`, `$top`, `$skip`, `$orderby`, and `$search`.
  - Folder scoping (Inbox, Junk Email, Drafts, custom folders).
  - HTML entity decoding, sanitization, and plain-text conversion.
  - Message attachment downloading and Base64 / Buffer decoding.
  - Message status manipulation (`markAsRead`, `deleteMessage`).
- **High-Level Polling Engine (`waitForEmail`)**:
  - Non-blocking async polling loop with customizable intervals, timeouts, and `receivedAfter` date thresholds.
  - Synchronous custom predicates for complex matching criteria.
  - Optional automatic mark-as-read upon successful match.
- **Intelligent Multi-Language OTP Extractor**:
  - Multi-tiered heuristic engine extracting 4–8 digit verification codes, hyphenated codes (`123-456`), and alphanumeric tokens.
  - Comprehensive keyword dictionary in English and Vietnamese (`verification code`, `security code`, `mã xác thực`, `mã OTP`, etc.).
  - False-positive filtering for copyright years (1900–2099), dates, and order numbers.
  - Confidence scoring (0.0 to 1.0) and context snippet generation.
- **Comprehensive CLI & Test Suite**:
  - Full-featured CLI with human-friendly and `--json` outputs.
  - 100% unit test coverage with mocked Graph APIs and live integration test runners.

---

## Installation & Setup

### Prerequisites

- **Node.js**: version `18.0.0` or higher (Node 20+ recommended).
- **Package Manager**: `npm`, `pnpm`, `yarn`, or `bun`.

### Installation

```bash
# Clone the repository
git clone https://github.com/your-org/graph-module.git
cd graph-module

# Install dependencies
npm install

# Build the project
npm run build
```

### TypeScript Configuration

Ensure your `tsconfig.json` targets `ES2022` or `ESNext` with `moduleResolution: "node16"` or `"bundler"`:

```json
{
  "compilerOptions": {
    "target": "ES2022",
    "module": "NodeNext",
    "moduleResolution": "NodeNext",
    "strict": true,
    "esModuleInterop": true,
    "skipLibCheck": true
  }
}
```

---

## Account Format Specification

The module uses a standardized pipe-delimited format for account credentials, widely used in multi-account automation and credential stores:

```text
email|password|refresh_token|client_id
```

An optional 5th field `authority` can also be specified:

```text
email|password|refresh_token|client_id|authority
```

### Field Breakdown

| Field Index | Field Name | Required | Description | Example |
| :--- | :--- | :--- | :--- | :--- |
| `0` | `email` | **Yes** | Hotmail / Outlook / Microsoft 365 email address | `john.doe@hotmail.com` |
| `1` | `password` | No | Account password (stored for convenience; not used in OAuth flow) | `P@ssw0rd123` |
| `2` | `refreshToken`| **Yes** | Valid OAuth 2.0 Refresh Token issued by Microsoft Identity | `M.C504_BL2.0.U.-C...` |
| `3` | `clientId` | **Yes** | Azure AD Application (Client) ID | `9e5f94bc-e8a4-4e73-b8be-63364c29d753` |
| `4` | `authority` | No | Tenant authority (defaults to `common` or `consumers`) | `common` |

### Account File Conventions

When reading accounts from a text file (e.g., `authenticated-test-account`):
- **Empty lines** are ignored automatically.
- **Comment lines** starting with `#` are ignored.
- Leading and trailing whitespace around each field is trimmed.

```text
# Development Hotmail Accounts
user1@hotmail.com|pwd123|M.C504_BL2.0.U.token1|9e5f94bc-e8a4-4e73-b8be-63364c29d753
user2@outlook.com|pwd456|M.C504_BL2.0.U.token2|9e5f94bc-e8a4-4e73-b8be-63364c29d753
```

---

## Quick Start Guide

### 1. Parsing Account Credentials

```typescript
import {
  parseAccountLine,
  parseAccountLines,
  parseAccountFile,
} from 'outlook-graph-module';

// 1. Parse a single account string
const accountStr = 'user@hotmail.com|secret_pwd|M.C504_BL2.0...|9e5f94bc-e8a4-4e73-b8be-63364c29d753';
const account = parseAccountLine(accountStr);
console.log(`Parsed account: ${account.email}, Client ID: ${account.clientId}`);

// 2. Parse multiple accounts from a string
const rawContent = `
# List of accounts
user1@hotmail.com|pwd1|token1|client_id_1
user2@outlook.com|pwd2|token2|client_id_2
`;
const accounts = parseAccountLines(rawContent);

// 3. Parse accounts directly from a file
const fileAccounts = await parseAccountFile('./authenticated-test-account');
console.log(`Loaded ${fileAccounts.length} accounts from file.`);
```

### 2. Fetching & Filtering Messages

```typescript
import { OutlookMailClient, parseAccountLine } from 'outlook-graph-module';

const account = parseAccountLine('user@hotmail.com|pwd|refresh_token|client_id');
const client = new OutlookMailClient();

// List the 10 most recent messages from Inbox
const messages = await client.getMessages(account, {
  top: 10,
  unreadOnly: false,
});

for (const msg of messages) {
  console.log(`[${msg.id}] ${msg.subject} from ${msg.from?.emailAddress.address}`);
}

// Fetch messages with specific filters
const filtered = await client.getMessages(account, {
  folder: 'inbox',
  unreadOnly: true,
  subjectContains: 'Security code',
  fromContains: 'microsoft.com',
  top: 5,
  orderBy: 'receivedDateTime desc',
  select: ['id', 'subject', 'from', 'receivedDateTime', 'bodyPreview'],
});
```

### 3. Reading Email Body & Attachments

```typescript
import { OutlookMailClient, cleanHtml } from 'outlook-graph-module';

const client = new OutlookMailClient();

// Fetch message by ID (includes complete body content)
const message = await client.getMessageById(account, 'AAMkAGI2...');

console.log('Subject:', message.subject);
console.log('Format :', message.body.contentType); // 'html' or 'text'

// Clean HTML to readable plain text
const plainText = cleanHtml(message.body.content);
console.log('Body Text:\n', plainText);

// Retrieve attachments if available
if (message.hasAttachments) {
  const attachments = await client.getMessageAttachments(account, message.id);
  for (const att of attachments) {
    console.log(`Attachment: ${att.name} (${att.size} bytes, type: ${att.contentType})`);
    if (att.data) {
      // att.data is a Node.js Buffer
      console.log(`Downloaded ${att.data.length} bytes into Buffer.`);
    }
  }
}
```

### 4. Polling & Waiting for Incoming Emails

Use `waitForEmail` to poll the mailbox until an email arriving after the start time matches your criteria.

```typescript
import { OutlookMailClient } from 'outlook-graph-module';

const client = new OutlookMailClient();

console.log('Triggering verification email on third-party service...');

// Poll for incoming email with a 60-second timeout
const incomingEmail = await client.waitForEmail(account, {
  subjectContains: 'Verification Code',
  fromContains: 'auth@service.com',
  timeoutMs: 60000,    // Max wait time (default: 60s)
  intervalMs: 3000,     // Polling interval (default: 3s)
  receivedAfter: new Date(Date.now() - 30000), // Buffer: 30s ago
  markAsReadAfterMatch: true, // Automatically mark as read once received
});

console.log(`Received email: ${incomingEmail.subject}`);
console.log(`Body preview: ${incomingEmail.bodyPreview}`);
```

### 5. Extracting OTP / Verification Codes

The built-in OTP extraction engine supports English and Vietnamese emails, multi-digit codes, hyphenated codes, and confidence ranking.

```typescript
import { OutlookMailClient, extractOtp, extractAllOtps } from 'outlook-graph-module';

const client = new OutlookMailClient();
const latestEmail = await client.getLatestMessage(account);

if (latestEmail) {
  // Option A: Extract single highest-confidence OTP from OutlookMessage
  const otpResult = client.extractOtp(latestEmail);

  if (otpResult) {
    console.log(`Found OTP Code : ${otpResult.code}`);
    console.log(`Confidence     : ${(otpResult.confidence * 100).toFixed(0)}%`);
    console.log(`Pattern Matched: ${otpResult.patternMatched}`);
    console.log(`Context Snippet: "${otpResult.contextSnippet}"`);
  }

  // Option B: Extract all potential candidates from raw text/HTML
  const allCandidates = extractAllOtps(latestEmail.body.content, {
    locale: 'auto', // 'en' | 'vi' | 'auto'
    preferredLength: 6, // Prioritize 6-digit codes
  });

  console.log(`Detected ${allCandidates.length} candidate(s):`, allCandidates);
}
```

### 6. Managing Messages (Mark as Read / Delete)

```typescript
// Mark an email as read
await client.markAsRead(account, message.id, true);

// Mark an email as unread
await client.markAsRead(account, message.id, false);

// Move message to Deleted Items folder
await client.deleteMessage(account, message.id);
```

---

## CLI Usage Guide

The module includes an executable CLI tool for rapid testing, mailbox inspection, OTP extraction, and batch validation across multiple accounts.

Run via `npm run cli`, `npx tsx ts/cli.ts`, or the installed binary `outlook-graph`:

```bash
npx tsx ts/cli.ts <command> [options]
```

### CLI Options & Flags

| Flag | Short | Type | Description | Default |
| :--- | :--- | :--- | :--- | :--- |
| `--account` | `-a` | string | Single account string (`email\|pwd\|refresh_token\|client_id`) | - |
| `--file` | `-f` | string | Path to accounts file | `authenticated-test-account` |
| `--top` | `-t` | number | Maximum number of emails to retrieve | `10` |
| `--folder` | | string | Target mail folder (e.g. `inbox`, `junkemail`, `drafts`) | `inbox` |
| `--unread` | `-u` | boolean | Filter unread emails only | `false` |
| `--subject` | `-s` | string | Filter emails by subject substring | - |
| `--from` | | string | Filter emails by sender email substring | - |
| `--id` | | string | Specific message ID for `read` command | - |
| `--format` | | string | Body format for output: `text` or `html` | `text` |
| `--timeout` | | number | Polling timeout in seconds for `wait` / `otp` | `60` |
| `--wait` | `-w` | boolean | Wait/poll for incoming email before extracting OTP | `false` |
| `--json` | | boolean | Output results as structured JSON | `false` |
| `--help` | `-h` | boolean | Display help and usage instructions | - |

---

### Command Reference & Examples

#### `auth` - Test Token Refresh

Tests the OAuth 2.0 token exchange against Microsoft Identity endpoints for a single account or the first account in a file.

```bash
# Using a single account string
npx tsx ts/cli.ts auth -a "user@hotmail.com|pwd|M.C504_BL2.0...|9e5f94bc-e8a4-4e73-b8be-63364c29d753"

# Using an account file
npx tsx ts/cli.ts auth --file ./authenticated-test-account

# Output response as JSON
npx tsx ts/cli.ts auth --file ./authenticated-test-account --json
```

**Output Example:**
```text
[AUTH] Refreshing OAuth token for user@hotmail.com...
[SUCCESS] Access token obtained successfully!
  Token Type : Bearer
  Expires In : 3600s
  Scope      : https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/Mail.ReadWrite offline_access
  Access Token Preview: EwB4A8l6BAAURiN86U4qf7wA...
```

---

#### `list` - List Inbox/Folder Messages

Fetches and displays messages from the mailbox with optional filtering.

```bash
# List top 5 messages
npx tsx ts/cli.ts list --file ./authenticated-test-account --top 5

# List only unread messages matching a subject
npx tsx ts/cli.ts list -f ./authenticated-test-account -u -s "Security"

# List messages in Junk Email folder as JSON
npx tsx ts/cli.ts list -f ./authenticated-test-account --folder junkemail --json
```

---

#### `read` - Read Email by ID

Fetches the complete message details, body content, and lists any attachments.

```bash
# Read email as cleaned plain text
npx tsx ts/cli.ts read -f ./authenticated-test-account --id "AAMkAGI2AAAU..."

# Read email in raw HTML format
npx tsx ts/cli.ts read -f ./authenticated-test-account --id "AAMkAGI2AAAU..." --format html
```

---

#### `latest` - Get Latest Email

Retrieves the single newest email matching optional filter criteria.

```bash
# Get the newest email in Inbox
npx tsx ts/cli.ts latest -f ./authenticated-test-account

# Get the newest unread email from GitHub
npx tsx ts/cli.ts latest -f ./authenticated-test-account --unread --from "github.com"
```

---

#### `wait` - Poll for Matching Email

Polls the mailbox every few seconds until a new matching email arrives or timeout occurs.

```bash
# Wait up to 60 seconds for an email with subject containing "Verification"
npx tsx ts/cli.ts wait -f ./authenticated-test-account --subject "Verification" --timeout 60

# Wait for an unread email from a specific sender
npx tsx ts/cli.ts wait -f ./authenticated-test-account --from "no-reply@auth.com" --unread
```

---

#### `otp` - Extract Verification Code

Fetches the latest email (or polls for incoming email with `-w`) and runs the multi-language OTP extractor.

```bash
# Extract OTP from latest email
npx tsx ts/cli.ts otp -f ./authenticated-test-account

# Wait for incoming verification email and extract OTP immediately
npx tsx ts/cli.ts otp -f ./authenticated-test-account -w --subject "Security code" --timeout 60

# Extract OTP from a specific message ID and output JSON
npx tsx ts/cli.ts otp -f ./authenticated-test-account --id "AAMkAGI2..." --json
```

---

#### `test-all` - Batch Account Validation

Iterates through all accounts in an account file, validates OAuth token refresh, and tests reading inbox messages for each.

```bash
# Test all accounts in default file (authenticated-test-account)
npx tsx ts/cli.ts test-all

# Test all accounts in a specific file
npx tsx ts/cli.ts test-all --file ./my-accounts.txt
```

---

## API Reference

### `OutlookMailClient`

The primary interface for mailbox operations.

```typescript
import { OutlookMailClient, OutlookTokenManager } from 'outlook-graph-module';

// Initialize with default options
const client = new OutlookMailClient();

// Or initialize with custom token manager / options
const client = new OutlookMailClient({
  authority: 'common',
  expiryBufferSec: 300,
  maxRetries: 3,
  onTokenRotated: async (oldToken, newToken, account) => {
    console.log(`Token rotated for ${account.email}`);
  },
});
```

#### Methods

- **`getMessages(account, options?): Promise<OutlookMessage[]>`**
  Retrieves messages matching the provided OData filter options.
- **`getLatestMessage(account, options?): Promise<OutlookMessage | null>`**
  Convenience helper that retrieves the single newest message.
- **`getMessageById(account, messageId, options?): Promise<OutlookMessage>`**
  Retrieves a full message including complete body content and recipient lists.
- **`getMessageAttachments(account, messageId): Promise<OutlookAttachment[]>`**
  Retrieves all attachments for a message, automatically decoding Base64 payloads into Node.js `Buffer` objects.
- **`markAsRead(account, messageId, isRead = true): Promise<void>`**
  Updates the read/unread status of a message.
- **`deleteMessage(account, messageId): Promise<void>`**
  Deletes a message (moves to Deleted Items).
- **`waitForEmail(account, options?): Promise<OutlookMessage>`**
  Polls the mailbox until an incoming email satisfies the given search parameters or predicate.
- **`extractOtp(input, options?): OtpResult | null`**
  Extracts the highest-confidence OTP code from an `OutlookMessage` or raw text/HTML string.

---

### `OutlookTokenManager`

Handles headless OAuth 2.0 token exchange, in-memory caching, rotation callbacks, and rate-limit backoff.

```typescript
import { OutlookTokenManager } from 'outlook-graph-module';

const tokenManager = new OutlookTokenManager({
  authority: 'common',
  expiryBufferSec: 300, // Refresh 5 minutes before expiration
  maxRetries: 3,
  initialRetryDelayMs: 1000,
  requestTimeoutMs: 15000,
  onTokenRotated: async (oldRt, newRt, account) => {
    // Persist new refresh token to database / file
  },
});

// Retrieve access token (cached if valid, otherwise refreshed)
const accessToken = await tokenManager.getAccessToken(account);

// Force refresh access token immediately
const tokenResponse = await tokenManager.refreshAccessToken(account);

// Cache inspection & invalidation
const cached = tokenManager.getCachedToken(account.email);
tokenManager.clearCache(account.email); // Clear specific account
tokenManager.clearCache();             // Clear all cached tokens
```

---

### Account Parser Functions

- **`parseAccountLine(line: string, options?: ParseAccountOptions): AccountCredentials`**
  Parses a single `email|password|refresh_token|client_id[|authority]` string.
- **`parseAccountLines(content: string, options?: ParseAccountOptions): AccountCredentials[]`**
  Parses multiline content, ignoring empty lines and comments.
- **`parseAccountFile(filePath: string, options?: ParseAccountOptions): Promise<AccountCredentials[]>`**
  Asynchronously reads a file from disk and returns parsed account credentials.

---

### OTP Extractor Functions

- **`extractOtp(input: OutlookMessage | string, options?: OtpExtractOptions): OtpResult | null`**
  Returns the candidate with the highest confidence score.
- **`extractAllOtps(input: OutlookMessage | string, options?: OtpExtractOptions): OtpResult[]`**
  Returns all detected OTP candidates sorted by confidence score descending.
- **`cleanHtml(html: string): string`**
  Strips HTML tags, styles, scripts, SVGs, and decodes HTML entities into normalized plain text.

---

### Query Builder Functions

- **`buildODataQuery(options?: GetMessagesOptions): string`**
  Constructs a standard OData query string (e.g. `?$filter=...&$top=10&$orderby=...`).
- **`escapeODataString(value: string): string`**
  Escapes single quotes for OData string literals (`O'Connor` -> `O''Connor`).
- **`getMailEndpoint(folder?: string): string`**
  Resolves the Microsoft Graph endpoint for the root mailbox or a specific folder.

---

### Error Hierarchy

All custom errors extend from the base `GraphModuleError`:

```text
GraphModuleError (base class)
├── AccountParseError     // Invalid account line format or file missing
├── TokenRefreshError     // OAuth 2.0 token exchange failure
├── GraphApiError         // Microsoft Graph API non-2xx response
├── RateLimitError        // HTTP 429 Too Many Requests
└── TimeoutError          // Polling timeout exceeded in waitForEmail
```

---

## TypeScript Types & Interfaces

```typescript
export interface AccountCredentials {
  email: string;
  password?: string;
  refreshToken: string;
  clientId: string;
  authority?: string;
}

export interface OutlookMessage {
  id: string;
  conversationId?: string;
  subject: string;
  bodyPreview?: string;
  body: {
    contentType: 'text' | 'html';
    content: string;
  };
  from?: {
    emailAddress: {
      name?: string;
      address: string;
    };
  };
  toRecipients?: Array<{
    emailAddress: {
      name?: string;
      address: string;
    };
  }>;
  receivedDateTime: string;
  sentDateTime?: string;
  hasAttachments: boolean;
  isRead: boolean;
  webLink?: string;
  raw?: Record<string, unknown>;
}

export interface OtpResult {
  code: string;
  digits: number;
  confidence: number;
  patternMatched: string;
  contextSnippet?: string;
  rawText: string;
}

export interface WaitForEmailOptions {
  subjectContains?: string;
  fromContains?: string;
  receivedAfter?: Date | string;
  folder?: string;
  unreadOnly?: boolean;
  timeoutMs?: number;
  intervalMs?: number;
  predicate?: (message: OutlookMessage) => boolean;
  markAsReadAfterMatch?: boolean;
}
```

---

## Testing

The project uses [Vitest](https://vitest.dev/) for high-speed unit and integration testing.

```bash
# Run unit tests (with mocked Graph API responses)
npm run test

# Run live integration tests against authenticated-test-account
npm run test:integration

# Run all tests (unit + integration)
npm run test:all

# Typecheck codebase
npm run typecheck
```
