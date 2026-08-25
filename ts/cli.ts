#!/usr/bin/env node

import { parseAccountLine, parseAccountFile } from './parser/account-parser.js';
import { OutlookMailClient } from './mail/mail-client.js';
import { OutlookTokenManager } from './auth/token-manager.js';
import { cleanHtml } from './otp/otp-extractor.js';
import type { AccountCredentials } from './types/account.js';

interface CliArgs {
  command: string;
  accountStr?: string;
  filePath?: string;
  top?: number;
  folder?: string;
  unreadOnly?: boolean;
  subject?: string;
  from?: string;
  messageId?: string;
  format?: 'text' | 'html';
  timeoutSec?: number;
  json?: boolean;
  wait?: boolean;
  help?: boolean;
}

function parseCliArgs(args: string[]): CliArgs {
  const parsed: CliArgs = {
    command: args[0] || 'help',
  };

  for (let i = 1; i < args.length; i++) {
    const arg = args[i];
    if (arg === '--account' || arg === '-a') {
      parsed.accountStr = args[++i];
    } else if (arg === '--file' || arg === '-f') {
      parsed.filePath = args[++i];
    } else if (arg === '--top' || arg === '-t') {
      parsed.top = parseInt(args[++i], 10);
    } else if (arg === '--folder') {
      parsed.folder = args[++i];
    } else if (arg === '--unread' || arg === '-u') {
      parsed.unreadOnly = true;
    } else if (arg === '--subject' || arg === '-s') {
      parsed.subject = args[++i];
    } else if (arg === '--from') {
      parsed.from = args[++i];
    } else if (arg === '--id') {
      parsed.messageId = args[++i];
    } else if (arg === '--format') {
      parsed.format = args[++i] as 'text' | 'html';
    } else if (arg === '--timeout') {
      parsed.timeoutSec = parseInt(args[++i], 10);
    } else if (arg === '--json') {
      parsed.json = true;
    } else if (arg === '--wait' || arg === '-w') {
      parsed.wait = true;
    } else if (arg === '--help' || arg === '-h') {
      parsed.help = true;
    }
  }

  return parsed;
}

function printUsage(): void {
  console.log(`
Microsoft Outlook Graph API CLI Tool

Usage:
  npx tsx ts/cli.ts <command> [options]

Commands:
  auth       Test OAuth2 token refresh for an account or accounts file
  list       List emails from inbox or folder
  read       Read full content of an email by message ID
  latest     Get the latest email matching optional filters
  wait       Poll and wait for a new incoming email matching criteria
  otp        Extract verification/OTP code from latest or incoming email
  test-all   Test token refresh and inbox listing on all accounts in file
  help       Show this help message

Options:
  -a, --account <str>   Account string in format: email|password|refresh_token|client_id
  -f, --file <path>     Path to accounts file (default for test-all: authenticated-test-account)
  -t, --top <num>       Number of emails to list (default: 10)
  --folder <name>       Mail folder (e.g. inbox, junkemail)
  -u, --unread          Filter unread messages only
  -s, --subject <str>   Filter subject containing substring
  --from <str>          Filter sender containing substring
  --id <msgId>          Specific message ID to read
  --format <text|html>  Body format to output (default: text)
  --timeout <sec>       Timeout in seconds for wait/polling (default: 60)
  -w, --wait            Wait for incoming email before extracting OTP
  --json                Output results in JSON format
  -h, --help            Show help

Examples:
  npx tsx ts/cli.ts test-all
  npx tsx ts/cli.ts auth --account "user@hotmail.com|pwd|token|client_id"
  npx tsx ts/cli.ts list --file authenticated-test-account --top 5
  npx tsx ts/cli.ts otp --file authenticated-test-account
  npx tsx ts/cli.ts wait --file authenticated-test-account --subject "Security code" --timeout 60
`);
}

async function resolveAccount(cliArgs: CliArgs): Promise<AccountCredentials> {
  if (cliArgs.accountStr) {
    return parseAccountLine(cliArgs.accountStr);
  }
  const filePath = cliArgs.filePath || 'authenticated-test-account';
  const accounts = await parseAccountFile(filePath);
  if (accounts.length === 0) {
    throw new Error(`No valid accounts found in file '${filePath}'`);
  }
  return accounts[0];
}

async function main(): Promise<void> {
  const cliArgs = parseCliArgs(process.argv.slice(2));

  if (cliArgs.help || cliArgs.command === 'help' || !cliArgs.command) {
    printUsage();
    return;
  }

  const client = new OutlookMailClient();

  try {
    switch (cliArgs.command) {
      case 'auth': {
        const account = await resolveAccount(cliArgs);
        console.log(`[AUTH] Refreshing OAuth token for ${account.email}...`);
        const tokenResp = await client.tokenManager.refreshAccessToken(account);
        if (cliArgs.json) {
          console.log(JSON.stringify(tokenResp, null, 2));
        } else {
          console.log(`[SUCCESS] Access token obtained successfully!`);
          console.log(`  Token Type : ${tokenResp.token_type}`);
          console.log(`  Expires In : ${tokenResp.expires_in}s`);
          console.log(`  Scope      : ${tokenResp.scope}`);
          console.log(`  Access Token Preview: ${tokenResp.access_token.slice(0, 30)}...`);
        }
        break;
      }

      case 'list': {
        const account = await resolveAccount(cliArgs);
        console.log(`[LIST] Fetching messages for ${account.email}...`);
        const messages = await client.getMessages(account, {
          top: cliArgs.top ?? 10,
          folder: cliArgs.folder,
          unreadOnly: cliArgs.unreadOnly,
          subjectContains: cliArgs.subject,
          fromContains: cliArgs.from,
        });

        if (cliArgs.json) {
          console.log(JSON.stringify(messages, null, 2));
        } else {
          console.log(`Found ${messages.length} message(s):\n`);
          messages.forEach((msg, idx) => {
            const sender = msg.from ? `${msg.from.emailAddress.name || ''} <${msg.from.emailAddress.address}>` : '(Unknown)';
            const readStatus = msg.isRead ? '[READ]' : '[UNREAD]';
            console.log(`${idx + 1}. ${readStatus} ${msg.subject}`);
            console.log(`   From: ${sender} | Date: ${msg.receivedDateTime}`);
            console.log(`   ID  : ${msg.id}`);
            console.log(`   Preview: ${msg.bodyPreview || '(No preview)'}\n`);
          });
        }
        break;
      }

      case 'latest': {
        const account = await resolveAccount(cliArgs);
        const msg = await client.getLatestMessage(account, {
          folder: cliArgs.folder,
          unreadOnly: cliArgs.unreadOnly,
          subjectContains: cliArgs.subject,
          fromContains: cliArgs.from,
        });

        if (!msg) {
          console.log('No messages found matching criteria.');
          return;
        }

        if (cliArgs.json) {
          console.log(JSON.stringify(msg, null, 2));
        } else {
          const sender = msg.from ? `${msg.from.emailAddress.name || ''} <${msg.from.emailAddress.address}>` : '(Unknown)';
          console.log(`Latest Message:`);
          console.log(`Subject : ${msg.subject}`);
          console.log(`From    : ${sender}`);
          console.log(`Date    : ${msg.receivedDateTime}`);
          console.log(`ID      : ${msg.id}`);
          console.log(`\n--- Body Preview ---`);
          console.log(msg.bodyPreview || cleanHtml(msg.body.content));
        }
        break;
      }

      case 'read': {
        const account = await resolveAccount(cliArgs);
        if (!cliArgs.messageId) {
          console.error('Error: --id <messageId> is required for read command.');
          process.exit(1);
        }
        const msg = await client.getMessageById(account, cliArgs.messageId);
        const attachments = msg.hasAttachments
          ? await client.getMessageAttachments(account, cliArgs.messageId)
          : [];

        if (cliArgs.json) {
          console.log(JSON.stringify({ message: msg, attachments }, null, 2));
        } else {
          const sender = msg.from ? `${msg.from.emailAddress.name || ''} <${msg.from.emailAddress.address}>` : '(Unknown)';
          console.log(`Subject : ${msg.subject}`);
          console.log(`From    : ${sender}`);
          console.log(`Date    : ${msg.receivedDateTime}`);
          console.log(`Attachments: ${attachments.length}`);
          attachments.forEach((att) => console.log(`  - ${att.name} (${att.size} bytes)`));
          console.log(`\n--- Body Content (${cliArgs.format || msg.body.contentType}) ---`);
          console.log(
            cliArgs.format === 'html'
              ? msg.body.content
              : cleanHtml(msg.body.content)
          );
        }
        break;
      }

      case 'wait': {
        const account = await resolveAccount(cliArgs);
        const timeoutMs = (cliArgs.timeoutSec ?? 60) * 1000;
        console.log(`[WAIT] Polling for incoming email on ${account.email} (timeout: ${cliArgs.timeoutSec ?? 60}s)...`);
        const msg = await client.waitForEmail(account, {
          subjectContains: cliArgs.subject,
          fromContains: cliArgs.from,
          timeoutMs,
          unreadOnly: cliArgs.unreadOnly,
        });

        if (cliArgs.json) {
          console.log(JSON.stringify(msg, null, 2));
        } else {
          console.log(`\n[MATCHED] New email arrived!`);
          console.log(`Subject : ${msg.subject}`);
          console.log(`From    : ${msg.from?.emailAddress.address || 'Unknown'}`);
          console.log(`Date    : ${msg.receivedDateTime}`);
          console.log(`Preview : ${msg.bodyPreview}`);
        }
        break;
      }

      case 'otp': {
        const account = await resolveAccount(cliArgs);
        let targetMsg: typeof client extends { getLatestMessage(...args: unknown[]): Promise<infer R> } ? R : null = null;

        if (cliArgs.messageId) {
          targetMsg = await client.getMessageById(account, cliArgs.messageId);
        } else if (cliArgs.wait) {
          const timeoutMs = (cliArgs.timeoutSec ?? 60) * 1000;
          console.log(`[OTP] Waiting for verification email on ${account.email} (timeout: ${cliArgs.timeoutSec ?? 60}s)...`);
          targetMsg = await client.waitForEmail(account, {
            subjectContains: cliArgs.subject,
            fromContains: cliArgs.from,
            timeoutMs,
          });
        } else {
          console.log(`[OTP] Fetching latest email on ${account.email}...`);
          targetMsg = await client.getLatestMessage(account, {
            subjectContains: cliArgs.subject,
            fromContains: cliArgs.from,
          });
        }

        if (!targetMsg) {
          console.log('No email found to extract OTP from.');
          return;
        }

        const otp = client.extractOtp(targetMsg);
        if (cliArgs.json) {
          console.log(JSON.stringify({ message: targetMsg, otp }, null, 2));
        } else {
          console.log(`\nEmail Subject : ${targetMsg.subject}`);
          console.log(`Email Date    : ${targetMsg.receivedDateTime}`);
          if (otp) {
            console.log(`\n[OTP DETECTED]`);
            console.log(`  Code       : ${otp.code}`);
            console.log(`  Confidence : ${(otp.confidence * 100).toFixed(0)}%`);
            console.log(`  Pattern    : ${otp.patternMatched}`);
            console.log(`  Context    : "${otp.contextSnippet}"`);
          } else {
            console.log(`\n[OTP] No verification code detected in this email.`);
          }
        }
        break;
      }

      case 'test-all': {
        const filePath = cliArgs.filePath || 'authenticated-test-account';
        const accounts = await parseAccountFile(filePath);
        console.log(`\n=== Batch Testing ${accounts.length} Accounts from '${filePath}' ===\n`);

        const results = [];
        for (let i = 0; i < accounts.length; i++) {
          const acc = accounts[i];
          process.stdout.write(`[${i + 1}/${accounts.length}] Testing ${acc.email}... `);
          try {
            const tokenResp = await client.tokenManager.refreshAccessToken(acc);
            const messages = await client.getMessages(acc, { top: 3 });
            process.stdout.write(`OK (Token valid, ${messages.length} message(s) in inbox)\n`);
            results.push({
              email: acc.email,
              status: 'success',
              messageCount: messages.length,
              tokenExpiresIn: tokenResp.expires_in,
            });
          } catch (err: unknown) {
            const msg = err instanceof Error ? err.message : String(err);
            process.stdout.write(`FAILED: ${msg}\n`);
            results.push({
              email: acc.email,
              status: 'failed',
              error: msg,
            });
          }
        }

        console.log(`\nSummary: ${results.filter((r) => r.status === 'success').length}/${accounts.length} passed.`);
        break;
      }

      default:
        console.error(`Unknown command: '${cliArgs.command}'. Run with --help for available commands.`);
        process.exit(1);
    }
  } catch (err: unknown) {
    const errorMsg = err instanceof Error ? err.message : String(err);
    console.error(`\n[ERROR] ${errorMsg}`);
    process.exit(1);
  }
}

main();
