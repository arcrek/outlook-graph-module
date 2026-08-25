# Microsoft Outlook Graph API Module

A high-performance, enterprise-grade **TypeScript/Node.js & Python** module for interacting with Microsoft Outlook and Hotmail mailboxes via the Microsoft Graph API v1.0. Built for account automation, inbox monitoring, verification workflows, and intelligent multi-language (EN/VI) OTP extraction using pipe-delimited account credentials.

[![TypeScript](https://img.shields.io/badge/TypeScript-5.7+-blue.svg)](https://www.typescriptlang.org/)
[![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)](https://www.python.org/)
[![Tests](https://img.shields.io/badge/Tests-100%25%20Passing-brightgreen.svg)]()
[![Zero External Runtime Dependencies](https://img.shields.io/badge/Dependencies-Zero%20Runtime%20Deps-orange.svg)]()
[![License](https://img.shields.io/badge/License-MIT-purple.svg)](./LICENSE)

Both implementations are functionally identical, built with zero external runtime dependencies, and provide the same CLI surface, account format, and OTP extraction heuristics.

---

## Documentation

| Guide | Covers |
| :--- | :--- |
| [docs/typescript.md](./docs/typescript.md) | TS/Node.js install, API reference, types, CLI usage, testing |
| [docs/python.md](./docs/python.md) | Python install, API reference, CLI usage, testing |
| [docs/architecture.md](./docs/architecture.md) | OAuth 2.0 flow, token rotation, rate-limit backoff, 4-tier OTP engine — applies to both implementations |

---

## Key Features

- **Zero Heavy Runtime Dependencies** — native `fetch`/`urllib.request`, no SDKs like `@microsoft/microsoft-graph-client` or `requests`.
- **Automated OAuth 2.0 Token Management** — headless refresh-token grant, in-memory caching (5-min safety buffer), rolling token rotation callbacks, in-flight de-duplication against token stampedes.
- **Built-in Resilience** — HTTP 429 exponential backoff with jitter, transparent 401 re-auth retry, transient 5xx recovery.
- **Rich Mailbox Operations** — full OData v4 (`$filter`, `$select`, `$top`, `$skip`, `$orderby`, `$search`), folder scoping, HTML sanitization, attachment decoding, mark-as-read/delete.
- **Non-blocking Polling (`waitForEmail` / `wait_for_email`)** — customizable interval/timeout/`receivedAfter`, custom predicates, optional auto mark-as-read.
- **Intelligent Multi-Language OTP Extractor** — 4-tier heuristic engine (prefix/suffix context, keyword proximity, false-positive filtering), EN + VI keyword dictionaries, confidence scoring.
- **CLI & Test Suite** — `auth`, `list`, `read`, `latest`, `wait`, `otp`, `test-all` commands with `--json` output; 100% unit test coverage plus live integration runners.

---

## Quickstart

### TypeScript / Node.js

```bash
npm install && npm run build
npx tsx ts/cli.ts otp -f ./authenticated-test-account
```
Full guide: [docs/typescript.md](./docs/typescript.md)

### Python

```bash
pip install -e .
python3 -m outlook_graph otp -f ./authenticated-test-account
```
Full guide: [docs/python.md](./docs/python.md)

---

## Account Format Specification

Both implementations share the same pipe-delimited account format:

```text
email|password|refresh_token|client_id[|authority]
```

| Field | Required | Description |
| :--- | :---: | :--- |
| `email` | Yes | Hotmail / Outlook / Microsoft 365 address |
| `password` | No | Stored for convenience; not used in OAuth flow |
| `refresh_token` | Yes | OAuth 2.0 refresh token issued by Microsoft Identity |
| `client_id` | Yes | Azure AD Application (Client) ID |
| `authority` | No | Tenant authority; defaults to `common`/`consumers` |

Empty lines and `#`-prefixed comment lines are ignored when reading account files (e.g. `authenticated-test-account`, gitignored by default).

See [docs/typescript.md](./docs/typescript.md#account-format-specification) or [docs/python.md](./docs/python.md#account-format-specification) for full parsing examples.

---

## Testing

```bash
# TypeScript (Vitest)
npm run test            # unit
npm run test:integration
npm run test:all

# Python (unittest)
python3 -m unittest discover -s tests/unit
python3 -m unittest discover -s tests/integration
```

---

## License

This project is licensed under the [MIT License](./LICENSE).
