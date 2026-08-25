import { readFile } from 'node:fs/promises';
import { AccountParseError } from '../types/errors.js';
import type { AccountCredentials, ParseAccountOptions } from '../types/account.js';

export function parseAccountLine(
  line: string,
  options?: ParseAccountOptions
): AccountCredentials {
  const shouldTrim = options?.trim ?? true;
  const raw = shouldTrim ? line.trim() : line;

  if (!raw) {
    throw new AccountParseError('Cannot parse empty account line', line);
  }

  // Format: email|password|refresh_token|client_id
  const parts = raw.split('|');

  if (parts.length < 4) {
    throw new AccountParseError(
      `Invalid account format: expected at least 4 pipe-delimited fields (email|password|refresh_token|client_id), received ${parts.length} field(s)`,
      line
    );
  }

  const email = (shouldTrim ? parts[0].trim() : parts[0]) || '';
  const password = shouldTrim ? parts[1].trim() : parts[1];
  const refreshToken = (shouldTrim ? parts[2].trim() : parts[2]) || '';
  const clientId = (shouldTrim ? parts[3].trim() : parts[3]) || '';
  const authority = (parts[4] ? (shouldTrim ? parts[4].trim() : parts[4]) : undefined) || options?.defaultAuthority;

  if (!email) {
    throw new AccountParseError('Account email is missing or empty', line);
  }
  if (!refreshToken) {
    throw new AccountParseError('Account refresh token is missing or empty', line);
  }
  if (!clientId) {
    throw new AccountParseError('Account client ID is missing or empty', line);
  }

  return {
    email,
    password: password || undefined,
    refreshToken,
    clientId,
    authority: authority || undefined,
  };
}

export function parseAccountLines(
  content: string,
  options?: ParseAccountOptions
): AccountCredentials[] {
  const ignoreEmpty = options?.ignoreEmptyLines ?? true;
  const ignoreComments = options?.ignoreComments ?? true;

  const lines = content.split(/\r?\n/);
  const results: AccountCredentials[] = [];

  for (let i = 0; i < lines.length; i++) {
    const rawLine = lines[i];
    const trimmed = rawLine.trim();

    if (ignoreEmpty && !trimmed) {
      continue;
    }
    if (ignoreComments && trimmed.startsWith('#')) {
      continue;
    }

    try {
      const parsed = parseAccountLine(rawLine, options);
      results.push(parsed);
    } catch (err) {
      if (err instanceof AccountParseError) {
        throw new AccountParseError(err.message, rawLine, i + 1);
      }
      throw err;
    }
  }

  return results;
}

export async function parseAccountFile(
  filePath: string,
  options?: ParseAccountOptions
): Promise<AccountCredentials[]> {
  try {
    const content = await readFile(filePath, 'utf-8');
    return parseAccountLines(content, options);
  } catch (err: unknown) {
    if (err instanceof AccountParseError) {
      throw err;
    }
    const message = err instanceof Error ? err.message : String(err);
    throw new AccountParseError(`Failed to read account file at '${filePath}': ${message}`);
  }
}
