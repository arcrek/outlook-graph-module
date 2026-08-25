# Architecture & Technical Design

This document details the architectural design, security considerations, protocol interactions, resilience mechanisms, and algorithmic heuristics implemented in the **Microsoft Outlook Graph API Module**.

---

## Table of Contents

1. [System Overview & Component Topology](#1-system-overview--component-topology)
2. [OAuth 2.0 Authentication & Token Lifecycle](#2-oauth-20-authentication--token-lifecycle)
   - [Refresh Token Grant Flow](#refresh-token-grant-flow)
   - [In-Memory Token Caching & Proactive Refresh](#in-memory-token-caching--proactive-refresh)
   - [Request De-duplication & Stampede Prevention](#request-de-duplication--stampede-prevention)
3. [Token Rotation Mechanics](#3-token-rotation-mechanics)
   - [Rolling Refresh Token Lifecycle](#rolling-refresh-token-lifecycle)
   - [The `onTokenRotated` Callback Architecture](#the-ontokenrotated-callback-architecture)
4. [Resilience & Rate Limiting Strategy](#4-resilience--rate-limiting-strategy)
   - [Microsoft Graph Throttling (HTTP 429) & `Retry-After`](#microsoft-graph-throttling-http-429--retry-after)
   - [Exponential Backoff with Full Jitter](#exponential-backoff-with-full-jitter)
   - [Transparent 401 Re-Authentication Recovery](#transparent-401-re-authentication-recovery)
   - [Transient 5xx Server Error Handling](#transient-5xx-server-error-handling)
5. [Mail Operations & OData Query Pipeline](#5-mail-operations--odata-query-pipeline)
   - [OData v4 Query Construction](#odata-v4-query-construction)
   - [Eventual Consistency & Header Management](#eventual-consistency--header-management)
   - [Non-blocking Polling Loop (`waitForEmail`)](#non-blocking-polling-loop-waitforemail)
6. [Intelligent OTP Extraction Engine](#6-intelligent-otp-extraction-engine)
   - [4-Tier Extraction Pipeline](#4-tier-extraction-pipeline)
   - [HTML Sanitization & Entity Decoding](#html-sanitization--entity-decoding)
   - [Multilingual Keyword Dictionaries (EN & VI)](#multilingual-keyword-dictionaries-en--vi)
   - [Confidence Scoring Algorithm & False-Positive Elimination](#confidence-scoring-algorithm--false-positive-elimination)
7. [Error Handling & Diagnostic Observability](#7-error-handling--diagnostic-observability)

---

## 1. System Overview & Component Topology

The module is engineered as a lean, zero-heavy-dependency client library. It bypasses monolithic SDKs in favor of direct HTTP calls to Microsoft Identity v2.0 and Microsoft Graph API v1.0 using modern native Node.js APIs (`fetch`, `AbortSignal`, `URLSearchParams`).

### Component Diagram

```
+-----------------------------------------------------------------------------------+
|                               Caller Application                                  |
+-----------------------------------------------------------------------------------+
       |                                      |                               |
       v                                      v                               v
+------------------+              +-----------------------+       +-----------------+
|  AccountParser   |              |   OutlookMailClient   |       |  OtpExtractor   |
| (pipe-delimited) |              +-----------------------+       +-----------------+
+------------------+                          |                            ^
       |                                      |                            |
       | creates                              | calls methods              | extracts
       v                                      v                            |
+--------------------+            +-----------------------+                |
| AccountCredentials | ---------> |    GraphHttpClient    |                |
+--------------------+            +-----------------------+                |
                                              |                            |
                             delegates auth   | queries API                |
                                              v                            |
                                  +-----------------------+                |
                                  |  OutlookTokenManager  |                |
                                  | - Cache (Map)         |                |
                                  | - In-flight dedup     |                |
                                  | - Rate limit backoff  |                |
                                  +-----------------------+                |
                                              |                            |
                                              | Bearer token               |
                                              v                            |
                       +---------------------------------------------+     |
                       |        Microsoft Cloud Infrastructure       |     |
                       |  - login.microsoftonline.com (OAuth 2.0)    |     |
                       |  - graph.microsoft.com/v1.0 (Mail API)      | ----+
                       +---------------------------------------------+
```

---

## 2. OAuth 2.0 Authentication & Token Lifecycle

Authentication in Microsoft Graph for automated personal accounts (Hotmail, Outlook.com, Live.com) and organizational accounts requires OAuth 2.0. Rather than interactive browser logins (Authorization Code Flow with PKCE) which cannot be automated headlessly, the module utilizes pre-authorized **OAuth 2.0 Refresh Tokens**.

### Refresh Token Grant Flow

```
OutlookMailClient         OutlookTokenManager              Microsoft Identity v2.0
      |                            |                                  |
      |-- getMessages(account) --->|                                  |
      |                            |-- Check in-memory cache          |
      |                            |   (Is token valid & unexpired?)  |
      |                            |                                  |
      |                            |-- [Cache Miss or Expiring Soon]  |
      |                            |-- POST /common/oauth2/v2.0/token |
      |                            |   grant_type=refresh_token       |
      |                            |   client_id=...                  |
      |                            |   refresh_token=...              |
      |                            |   scope=Mail.Read...             |
      |                            |--------------------------------->|
      |                            |                                  |
      |                            |<-- HTTP 200 OK ------------------|
      |                            |    { access_token, expires_in,   |
      |                            |      refresh_token (optional) }  |
      |                            |                                  |
      |                            |-- Store in cache                 |
      |                            |   expiresAt = now + expires_in   |
      |                            |-- Return access_token            |
      |<-- Bearer Access Token ----|                                  |
```

### In-Memory Token Caching & Proactive Refresh

Tokens are cached per account in a normalized `Map<string, TokenCacheEntry>` keyed by lowercase email.

- **Expiration Safety Window (`expiryBufferSec`)**: Default is **300 seconds (5 minutes)**.
- **Validity Rule**:
  $$\text{isValid} = \text{Date.now}() < (\text{cached.expiresAt} - \text{expiryBufferMs})$$
- By proactively refreshing before absolute token expiration, the client completely prevents race conditions where a token expires in-transit between our client and Microsoft Graph.

### Request De-duplication & Stampede Prevention

In concurrent architectures (e.g., polling multiple mailboxes or issuing simultaneous requests for the same account), multiple asynchronous workers might discover an expired token simultaneously.

Without de-duplication, this causes a **token stampede** where 10 concurrent requests fire 10 simultaneous OAuth refresh calls. Microsoft Identity throttles or invalidates refresh tokens under rapid simultaneous usage.

The module implements **In-Flight Promise Sharing**:

```typescript
// Inside OutlookTokenManager
private readonly inFlightRefreshes = new Map<string, Promise<TokenResponse>>();

public async refreshAccessToken(account: AccountCredentials): Promise<TokenResponse> {
  const key = account.email.toLowerCase();
  
  // Return existing in-flight promise if one is already executing
  const existingInFlight = this.inFlightRefreshes.get(key);
  if (existingInFlight) {
    return existingInFlight;
  }

  const refreshPromise = this.executeRefresh(account);
  this.inFlightRefreshes.set(key, refreshPromise);

  try {
    return await refreshPromise;
  } finally {
    this.inFlightRefreshes.delete(key);
  }
}
```

---

## 3. Token Rotation Mechanics

### Rolling Refresh Token Lifecycle

Microsoft identity platforms frequently employ **Rolling Refresh Tokens** (Token Rotation). When exchanging an old refresh token for a new access token, the identity provider may return a brand-new `refresh_token` in the response payload and invalidate the previous one after a short grace period.

If an application ignores the newly issued `refresh_token` and continues using the old one on subsequent runs, the account will eventually become permanently unauthenticated.

### The `onTokenRotated` Callback Architecture

To ensure persistence layers (PostgreSQL, Redis, flat files) can immediately update stored credentials without coupling the token manager to a specific database:

```
+---------------------+      New refresh_token returned      +----------------------+
| OutlookTokenManager | -----------------------------------> | onTokenRotated Hook  |
+---------------------+                                      +----------------------+
                                                                        |
                                                                        v
                                                             +----------------------+
                                                             | External Persistence |
                                                             | (DB / Encrypted File)|
                                                             +----------------------+
```

1. When a new `refresh_token` is present and distinct from `account.refreshToken`:
   ```typescript
   account.refreshToken = tokenResponse.refresh_token;
   if (this.onTokenRotated) {
     await this.onTokenRotated(oldRefreshToken, tokenResponse.refresh_token, account);
   }
   ```
2. The in-memory `AccountCredentials` object is updated in-place immediately.
3. The asynchronous callback is invoked safely within a `try-catch` wrapper so persistence failures never break the active HTTP request.

---

## 4. Resilience & Rate Limiting Strategy

### Microsoft Graph Throttling (HTTP 429) & `Retry-After`

Microsoft Graph enforces per-app, per-mailbox, and global tenant rate limits. When limits are exceeded, Graph returns `HTTP 429 Too Many Requests` along with a standard `Retry-After` header specifying the mandatory backoff duration in seconds.

### Exponential Backoff with Full Jitter

When a 429 is encountered, the client parses the `Retry-After` header. If absent or invalid, it falls back to an exponential backoff formula with randomized jitter to prevent cluster synchrony:

$$t_{\text{wait}} = \begin{cases} 
(\text{Retry-After} \times 1000) + \text{rand}(0, 200)\text{ms} & \text{if Header exists} \\ 
(\text{initialDelay} \times 2^{\text{attempt}-1}) + \text{rand}(0, 200)\text{ms} & \text{otherwise} 
\end{cases}$$

```typescript
// Retry loop logic in GraphHttpClient & OutlookTokenManager
if (res.status === 429) {
  const retryAfterHeader = res.headers.get('Retry-After');
  const retryAfterSec = retryAfterHeader ? parseInt(retryAfterHeader, 10) : NaN;
  const retryDelayMs = !isNaN(retryAfterSec) && retryAfterSec > 0
    ? retryAfterSec * 1000
    : delay;

  if (attempt > maxRetries) {
    throw new RateLimitError('Rate limit exceeded', retryDelayMs, 429);
  }

  const { promise, resolve } = Promise.withResolvers<void>();
  setTimeout(resolve, retryDelayMs + Math.random() * 200);
  await promise;

  delay *= 2;
  continue;
}
```

### Transparent 401 Re-Authentication Recovery

If an access token is revoked prematurely on Microsoft's servers, Graph API returns `HTTP 401 Unauthorized`. Instead of failing immediately:

1. `GraphHttpClient` catches the initial `401`.
2. It sets `didRetryAuth = true`.
3. It calls `tokenManager.getAccessToken(account, /* forceRefresh */ true)`.
4. It repeats the request with the fresh token.
5. If the request fails a second time with 401, it throws `GraphApiError`.

### Transient 5xx Server Error Handling

Microsoft Graph occasionally experiences temporary downstream gateway timeouts (`503 Service Unavailable`, `504 Gateway Timeout`, `500 Internal Server Error`).

The client retries up to `maxRetries` (default: 3) with progressive exponential backoff before throwing a descriptive `GraphApiError`.

---

## 5. Mail Operations & OData Query Pipeline

### OData v4 Query Construction

Microsoft Graph uses OData v4 query conventions. The `buildODataQuery` utility transforms strongly typed TypeScript query options into compliant URL query strings:

| TypeScript Option | OData Output | Purpose |
| :--- | :--- | :--- |
| `unreadOnly: true` | `$filter=isRead eq false` | Filters for unread messages only |
| `subjectContains: "foo"` | `$filter=contains(subject, 'foo')` | Substring match on subject line |
| `fromContains: "bar"` | `$filter=contains(from/emailAddress/address, 'bar')` | Substring match on sender address |
| `receivedAfter: Date` | `$filter=receivedDateTime ge 2026-08-25T10:00:00.000Z` | Threshold filter for polling |
| `top: 10` | `$top=10` | Page size limiter (clamped 1–1000) |
| `skip: 20` | `$skip=20` | Offset pagination |
| `select: ['id', 'subject']` | `$select=id,subject` | Projection (reduces bandwidth) |
| `orderBy: 'receivedDateTime desc'` | `$orderby=receivedDateTime desc` | Default ordering by newest first |
| `search: "code"` | `$search="code"` | Full-text search (KQL) |

#### OData String Literal Escaping

To prevent query injection or malformed OData syntax when filtering names with quotes (e.g. `O'Connor`), single quotes are doubled:
```typescript
export function escapeODataString(value: string): string {
  return value.replace(/'/g, "''");
}
```

### Eventual Consistency & Header Management

When performing `$search` queries or deep mailbox filters in Microsoft Graph, Graph requires the `ConsistencyLevel: eventual` header. `GraphHttpClient` automatically injects this header when `search` options are detected.

### Non-blocking Polling Loop (`waitForEmail`)

The `waitForEmail` method enables automated end-to-end verification workflows:

```
[Start Polling]
       |
       v
[Calculate receivedAfter = now - buffer]
       |
       +---> [Fetch messages receivedDateTime >= receivedAfter]
       |                      |
       |                      v
       |             [Check Predicates]
       |             - subjectContains?
       |             - fromContains?
       |             - custom predicate(msg)?
       |                      |
       |            +---------+---------+
       |            | Match Found       | No Match
       |            v                   v
       |     [Optional: markAsRead]   [Elapsed > timeoutMs?]
       |            |                   |          |
       |            v                   v Yes      v No
       |     [Return Message]     [Throw Timeout]  [Wait intervalMs (3s)]
       |                                                   |
       +---------------------------------------------------+
```

---

## 6. Intelligent OTP Extraction Engine

Extracting One-Time Passwords (OTPs) from diverse emails sent by thousands of different identity providers (Microsoft, Google, GitHub, Steam, banks, crypto exchanges) is challenging due to inconsistent layouts, HTML markup noise, copyright dates, and multi-language phrasing.

### 4-Tier Extraction Pipeline

```
                              Raw Input (HTML / Text / OutlookMessage)
                                                 │
                                                 ▼
                                     ┌───────────────────────┐
                                     │   cleanHtml Engine    │
                                     │ - Strip <style>/<svg> │
                                     │ - Strip all HTML tags │
                                     │ - Decode &entity;     │
                                     └───────────────────────┘
                                                 │
                                           Cleaned Text
                                                 │
         ┌───────────────────────────────────────┴───────────────────────────────────────┐
         ▼                                       ▼                                       ▼
  Tier 1: Context Prefix                  Tier 2: Context Suffix                 Tier 3: Keyword Proximity
  "code is 592814"                        "592814 is your code"                  Proximity to EN/VI keywords
  (Confidence: 0.98)                      (Confidence: 0.95)                     (Confidence: 0.65 - 0.99)
         │                                       │                                       │
         └───────────────────────────────────────┬───────────────────────────────────────┘
                                                 │
                                                 ▼
                                  ┌─────────────────────────────┐
                                  │ Tier 4: False Positive Filter│
                                  │ - Exclude copyright years   │
                                  │ - Exclude dates (19xx, 20xx)│
                                  └─────────────────────────────┘
                                                 │
                                                 ▼
                                  ┌─────────────────────────────┐
                                  │ Confidence Ranking & Sort   │
                                  │ Return top candidate        │
                                  └─────────────────────────────┘
```

---

### HTML Sanitization & Entity Decoding

Before regex analysis, emails are sanitized:
1. Complete removal of `<style>`, `<script>`, and `<svg>` blocks including inner text.
2. Stripping of all HTML tags (`<[^>]+>`).
3. Decoding of named entities (`&nbsp;`, `&amp;`, `&quot;`, `&apos;`, `&#39;`).
4. Removal of Zero-Width Non-Joiners (`&zwnj;`, `&zwj;`) often injected by spammers or template builders.
5. Decoding of arbitrary decimal (`&#123;`) and hexadecimal (`&#x1F;`) character references.
6. Normalization of whitespace.

---

### Multilingual Keyword Dictionaries (EN & VI)

The engine features extensive built-in dictionaries for English and Vietnamese transactional templates:

#### English Keywords (`KEYWORDS_EN`)
- `verification code`, `security code`, `confirmation code`
- `one-time password`, `one-time passcode`, `otp`
- `passcode`, `secret code`, `access code`, `login code`
- `authorization code`, `auth code`, `validation code`, `pin code`
- `temporary password`

#### Vietnamese Keywords (`KEYWORDS_VI`)
- `mã xác thực`, `mã xác nhận`, `mã otp`, `mã bảo mật`
- `mã kiểm tra`, `mã xác minh`, `mật khẩu một lần`
- `mã kích hoạt`, `mã an toàn`, `mã giao dịch`
- `mã đăng nhập`, `mã số bí mật`

---

### Confidence Scoring Algorithm & False-Positive Elimination

Each potential candidate code is evaluated and assigned a confidence score:

| Extraction Pattern | Description | Base Confidence |
| :--- | :--- | :--- |
| **`context_prefix`** | Code immediately preceded by keyword + connector (`"verification code: 123456"`, `"mã xác thực là: 123456"`) | **0.98** |
| **`context_suffix`** | Code immediately followed by confirmation phrase (`"123456 is your code"`, `"123456 là mã xác nhận"`) | **0.95** |
| **`numeric_6_proximate`** | 6-digit number $\le 60$ characters from an OTP keyword | **0.99** |
| **`numeric_hyphenated`** | 6-digit hyphenated code (`123-456`) $\le 60$ characters from keyword | **0.99** |
| **`numeric_8_proximate`** | 8-digit number $\le 60$ characters from keyword | **0.95** |
| **`numeric_4_proximate`** | 4-digit number $\le 60$ characters from keyword | **0.85** |
| **`numeric_isolated`** | Number isolated without nearby keyword ($> 300$ chars away) | **0.25 – 0.40** |

#### False Positive Elimination:
- **Copyright Years**: 4-digit numbers starting with `19` or `20` preceded or followed by `Copyright`, `©`, `(c)`, or `All rights reserved` are filtered out.
- **Hyphen Normalization**: Formats like `491-823` or `491 823` are normalized to `491823` while retaining the original raw string in `OtpResult.rawText`.

---

## 7. Error Handling & Diagnostic Observability

All exceptions thrown by the module derive from `GraphModuleError`, providing clear classification and diagnostic context:

```
                      GraphModuleError
                      (code, message)
                             │
       ┌─────────────────────┼─────────────────────┬─────────────────────┐
       ▼                     ▼                     ▼                     ▼
AccountParseError    TokenRefreshError       GraphApiError         RateLimitError
(rawLine, lineNum)   (statusCode, response) (statusCode, code)   (retryAfterMs)
                                                   │
                                                   ▼
                                              TimeoutError
                                            (polling timeout)
```

### Error Properties Reference

```typescript
try {
  await client.getMessages(account);
} catch (err) {
  if (err instanceof RateLimitError) {
    console.warn(`Rate limited. Backing off for ${err.retryAfterMs}ms`);
  } else if (err instanceof TokenRefreshError) {
    console.error(`Auth failed with HTTP ${err.statusCode}:`, err.errorResponse);
  } else if (err instanceof GraphApiError) {
    console.error(`Graph returned ${err.statusCode} [${err.graphErrorCode}]: ${err.message}`);
  } else if (err instanceof AccountParseError) {
    console.error(`Invalid account format on line ${err.lineNumber}: "${err.rawLine}"`);
  }
}
```
