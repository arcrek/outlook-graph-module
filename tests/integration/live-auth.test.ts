import { describe, it, expect } from 'vitest';
import { resolve } from 'node:path';
import { parseAccountFile } from '../../ts/parser/account-parser.js';
import { OutlookTokenManager } from '../../ts/auth/token-manager.js';

describe('live-auth integration', () => {
  it('should refresh access token for all accounts in authenticated-test-account', async () => {
    const filePath = resolve(process.cwd(), 'authenticated-test-account');
    const accounts = await parseAccountFile(filePath);

    expect(accounts.length).toBeGreaterThanOrEqual(1);

    const tokenManager = new OutlookTokenManager();

    for (const acc of accounts) {
      const response = await tokenManager.refreshAccessToken(acc);
      expect(response.access_token).toBeDefined();
      expect(response.access_token.length).toBeGreaterThan(50);
      expect(response.expires_in).toBeGreaterThan(0);
      expect(response.token_type).toBe('Bearer');
    }
  });
});
