import { describe, it, expect } from 'vitest';
import { resolve } from 'node:path';
import { parseAccountFile } from '../../ts/parser/account-parser.js';
import { OutlookMailClient } from '../../ts/mail/mail-client.js';

describe('live-mail integration', () => {
  it('should list messages from real Hotmail inbox', async () => {
    const filePath = resolve(process.cwd(), 'authenticated-test-account');
    const accounts = await parseAccountFile(filePath);
    expect(accounts.length).toBeGreaterThan(0);

    const client = new OutlookMailClient();
    const primaryAccount = accounts[0];

    const messages = await client.getMessages(primaryAccount, { top: 5 });
    expect(Array.isArray(messages)).toBe(true);

    if (messages.length > 0) {
      const first = messages[0];
      expect(first.id).toBeDefined();
      expect(first.subject).toBeDefined();

      const fullMessage = await client.getMessageById(primaryAccount, first.id);
      expect(fullMessage.id).toBe(first.id);
      expect(fullMessage.body).toBeDefined();
    }
  });
});
