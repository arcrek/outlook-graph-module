import { describe, it, expect } from 'vitest';
import {
  parseAccountLine,
  parseAccountLines,
  parseAccountFile,
} from '../../ts/parser/account-parser.js';
import { AccountParseError } from '../../ts/types/errors.js';
import { resolve } from 'node:path';
import { existsSync } from 'node:fs';

describe('account-parser', () => {
  it('should parse a valid 4-part account line', () => {
    const line =
      'user@hotmail.com|password123|M.C557_BAY.token...|9e5f94bc-e8a4-4e73-b8be-63364c29d753';
    const acc = parseAccountLine(line);

    expect(acc.email).toBe('user@hotmail.com');
    expect(acc.password).toBe('password123');
    expect(acc.refreshToken).toBe('M.C557_BAY.token...');
    expect(acc.clientId).toBe('9e5f94bc-e8a4-4e73-b8be-63364c29d753');
    expect(acc.authority).toBeUndefined();
  });

  it('should parse optional authority in 5-part account line', () => {
    const line =
      'user@hotmail.com|password123|token|client-id|consumers';
    const acc = parseAccountLine(line);

    expect(acc.authority).toBe('consumers');
  });
  it('should parse recovery email in 5-part account line when 5th field contains @', () => {
    const line =
      'user@hotmail.com|password123|token|client-id|user@recovery.com';
    const acc = parseAccountLine(line);

    expect(acc.recoveryEmail).toBe('user@recovery.com');
    expect(acc.authority).toBeUndefined();
  });

  it('should parse 6-part account line with recovery email and authority', () => {
    const line =
      'user@hotmail.com|password123|token|client-id|user@recovery.com|consumers';
    const acc = parseAccountLine(line);

    expect(acc.recoveryEmail).toBe('user@recovery.com');
    expect(acc.authority).toBe('consumers');
  });

  it('should parse 6-part account line with authority and recovery email', () => {
    const line =
      'user@hotmail.com|password123|token|client-id|consumers|user@recovery.com';
    const acc = parseAccountLine(line);

    expect(acc.recoveryEmail).toBe('user@recovery.com');
    expect(acc.authority).toBe('consumers');
  });

  it('should throw AccountParseError for empty or whitespace-only lines', () => {
    expect(() => parseAccountLine('')).toThrow(AccountParseError);
    expect(() => parseAccountLine('   ')).toThrow(AccountParseError);
  });

  it('should throw AccountParseError for lines with fewer than 4 fields', () => {
    expect(() => parseAccountLine('user@hotmail.com|pass')).toThrow(
      /Invalid account format/
    );
    expect(() =>
      parseAccountLine('user@hotmail.com|pass|only_three_parts')
    ).toThrow(/Invalid account format/);
  });

  it('should throw AccountParseError if required fields are missing', () => {
    expect(() =>
      parseAccountLine('|pass|refresh_token|client_id')
    ).toThrow(/email is missing/);
    expect(() =>
      parseAccountLine('user@hotmail.com|pass||client_id')
    ).toThrow(/refresh token is missing/);
    expect(() =>
      parseAccountLine('user@hotmail.com|pass|refresh_token|')
    ).toThrow(/client ID is missing/);
  });

  it('should parse multi-line account content ignoring comments and blank lines', () => {
    const content = `
# Hotmail Test Accounts
user1@hotmail.com|p1|rt1|cid1

# Secondary account
user2@hotmail.com|p2|rt2|cid2
`;

    const accounts = parseAccountLines(content);
    expect(accounts).toHaveLength(2);
    expect(accounts[0].email).toBe('user1@hotmail.com');
    expect(accounts[1].email).toBe('user2@hotmail.com');
  });

  it('should parse test account file correctly', async () => {
    const filePath = resolve(process.cwd(), 'authenticated-test-account');
    if (!existsSync(filePath)) return; // gitignored local fixture; skip if absent

    const accounts = await parseAccountFile(filePath);

    expect(accounts.length).toBeGreaterThanOrEqual(1);
    expect(accounts[0].email).toBeTruthy();
    expect(accounts[0].clientId).toBeTruthy();
  });
});
