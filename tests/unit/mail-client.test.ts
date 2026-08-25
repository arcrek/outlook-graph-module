import { describe, it, expect, vi, beforeEach } from 'vitest';
import { OutlookMailClient } from '../../ts/mail/mail-client.js';
import { OutlookTokenManager } from '../../ts/auth/token-manager.js';
import type { AccountCredentials } from '../../ts/types/account.js';

describe('mail-client', () => {
  const mockAccount: AccountCredentials = {
    email: 'user@hotmail.com',
    refreshToken: 'mock-rt',
    clientId: 'mock-cid',
  };

  let tokenManager: OutlookTokenManager;
  let mailClient: OutlookMailClient;

  beforeEach(() => {
    vi.restoreAllMocks();
    tokenManager = new OutlookTokenManager();
    vi.spyOn(tokenManager, 'getAccessToken').mockResolvedValue('mock-token');
    mailClient = new OutlookMailClient(tokenManager);
  });

  it('should fetch messages and map to OutlookMessage objects', async () => {
    const mockGraphResponse = {
      value: [
        {
          id: 'msg-1',
          subject: 'Your verification code is 849201',
          bodyPreview: 'Please enter 849201 to verify your account.',
          body: { contentType: 'html', content: '<p>Please enter <b>849201</b></p>' },
          from: { emailAddress: { name: 'Discord', address: 'noreply@discord.com' } },
          receivedDateTime: '2026-08-25T14:30:00Z',
          hasAttachments: false,
          isRead: false,
        },
      ],
    };

    vi.spyOn(globalThis, 'fetch').mockResolvedValueOnce(
      new Response(JSON.stringify(mockGraphResponse), {
        status: 200,
        headers: { 'Content-Type': 'application/json' },
      })
    );

    const messages = await mailClient.getMessages(mockAccount, { top: 1 });
    expect(messages).toHaveLength(1);
    expect(messages[0].id).toBe('msg-1');
    expect(messages[0].subject).toBe('Your verification code is 849201');
    expect(messages[0].from?.emailAddress.address).toBe('noreply@discord.com');
    expect(messages[0].isRead).toBe(false);
  });

  it('should fetch single message by ID', async () => {
    const mockMessage = {
      id: 'msg-abc',
      subject: 'Security Alert',
      body: { contentType: 'text', content: 'Unusual sign-in detected' },
      receivedDateTime: '2026-08-25T10:00:00Z',
      hasAttachments: false,
      isRead: true,
    };

    vi.spyOn(globalThis, 'fetch').mockResolvedValueOnce(
      new Response(JSON.stringify(mockMessage), {
        status: 200,
        headers: { 'Content-Type': 'application/json' },
      })
    );

    const message = await mailClient.getMessageById(mockAccount, 'msg-abc');
    expect(message.id).toBe('msg-abc');
    expect(message.subject).toBe('Security Alert');
  });

  it('should fetch and decode base64 attachments', async () => {
    const mockAttachmentsResponse = {
      value: [
        {
          id: 'att-1',
          name: 'invoice.pdf',
          contentType: 'application/pdf',
          size: 1024,
          isInline: false,
          contentBytes: Buffer.from('PDF Content').toString('base64'),
        },
      ],
    };

    vi.spyOn(globalThis, 'fetch').mockResolvedValueOnce(
      new Response(JSON.stringify(mockAttachmentsResponse), {
        status: 200,
        headers: { 'Content-Type': 'application/json' },
      })
    );

    const attachments = await mailClient.getMessageAttachments(
      mockAccount,
      'msg-1'
    );
    expect(attachments).toHaveLength(1);
    expect(attachments[0].name).toBe('invoice.pdf');
    expect(attachments[0].data?.toString('utf-8')).toBe('PDF Content');
  });

  it('should mark message as read', async () => {
    const fetchSpy = vi.spyOn(globalThis, 'fetch').mockResolvedValueOnce(
      new Response(null, { status: 204 })
    );

    await mailClient.markAsRead(mockAccount, 'msg-1', true);

    expect(fetchSpy).toHaveBeenCalledWith(
      'https://graph.microsoft.com/v1.0/me/messages/msg-1',
      expect.objectContaining({
        method: 'PATCH',
        body: JSON.stringify({ isRead: true }),
      })
    );
  });
});
