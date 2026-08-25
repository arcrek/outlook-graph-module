import { OutlookTokenManager } from '../auth/token-manager.js';
import { GraphHttpClient } from '../http/graph-client.js';
import { buildODataQuery, getMailEndpoint } from '../http/query-builder.js';
import { extractOtp } from '../otp/otp-extractor.js';
import { TimeoutError, TokenRefreshError, GraphApiError } from '../types/errors.js';
import type { AccountCredentials } from '../types/account.js';
import type { TokenManagerOptions } from '../types/auth.js';
import type {
  OutlookMessage,
  OutlookAttachment,
  GetMessagesOptions,
  WaitForEmailOptions,
} from '../types/mail.js';
import type { OtpExtractOptions, OtpResult } from '../types/otp.js';

export class OutlookMailClient {
  public readonly tokenManager: OutlookTokenManager;
  public readonly httpClient: GraphHttpClient;

  constructor(tokenManagerOrOptions?: OutlookTokenManager | TokenManagerOptions) {
    if (tokenManagerOrOptions instanceof OutlookTokenManager) {
      this.tokenManager = tokenManagerOrOptions;
    } else {
      this.tokenManager = new OutlookTokenManager(tokenManagerOrOptions);
    }
    this.httpClient = new GraphHttpClient(this.tokenManager);
  }

  public async getMessages(
    account: AccountCredentials,
    options?: GetMessagesOptions
  ): Promise<OutlookMessage[]> {
    const endpoint = getMailEndpoint(options?.folder);
    const queryString = buildODataQuery(options);
    const url = `${endpoint}${queryString}`;

    const res = await this.httpClient.get<{ value: Record<string, unknown>[] }>(
      account,
      url,
      { searchConsistencyEventual: Boolean(options?.search) }
    );

    const rawList = res.value || [];
    return rawList.map((raw) => this.mapToOutlookMessage(raw));
  }

  public async getLatestMessage(
    account: AccountCredentials,
    options?: Omit<GetMessagesOptions, 'top'>
  ): Promise<OutlookMessage | null> {
    const messages = await this.getMessages(account, {
      ...options,
      top: 1,
    });
    return messages[0] || null;
  }

  public async getMessageById(
    account: AccountCredentials,
    messageId: string,
    options?: { select?: string[] }
  ): Promise<OutlookMessage> {
    const selectQuery = options?.select?.length
      ? `?$select=${options.select.join(',')}`
      : '';
    const url = `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(
      messageId
    )}${selectQuery}`;

    const raw = await this.httpClient.get<Record<string, unknown>>(account, url);
    return this.mapToOutlookMessage(raw);
  }

  public async getMessageAttachments(
    account: AccountCredentials,
    messageId: string
  ): Promise<OutlookAttachment[]> {
    const url = `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(
      messageId
    )}/attachments`;

    const res = await this.httpClient.get<{ value: Record<string, unknown>[] }>(
      account,
      url
    );

    const rawList = res.value || [];
    return rawList.map((item) => {
      const contentBytes = item['contentBytes'] as string | undefined;
      const data = contentBytes ? Buffer.from(contentBytes, 'base64') : undefined;

      return {
        id: (item['id'] as string) || '',
        name: (item['name'] as string) || '',
        contentType: (item['contentType'] as string) || 'application/octet-stream',
        size: (item['size'] as number) || 0,
        isInline: Boolean(item['isInline']),
        contentBytes,
        data,
      };
    });
  }

  public async markAsRead(
    account: AccountCredentials,
    messageId: string,
    isRead = true
  ): Promise<void> {
    const url = `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(
      messageId
    )}`;
    await this.httpClient.patch(account, url, { isRead });
  }

  public async deleteMessage(
    account: AccountCredentials,
    messageId: string
  ): Promise<void> {
    const url = `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(
      messageId
    )}`;
    await this.httpClient.delete(account, url);
  }

  public async waitForEmail(
    account: AccountCredentials,
    options?: WaitForEmailOptions
  ): Promise<OutlookMessage> {
    const timeoutMs = options?.timeoutMs ?? 60000;
    const intervalMs = options?.intervalMs ?? 3000;
    const receivedAfter = options?.receivedAfter ?? new Date(Date.now() - 60000); // 1 min buffer by default
    const startTime = Date.now();

    while (true) {
      const elapsed = Date.now() - startTime;
      if (elapsed > timeoutMs) {
        throw new TimeoutError(
          `Timed out after ${timeoutMs}ms waiting for matching email on account ${account.email}`
        );
      }

      try {
        const messages = await this.getMessages(account, {
          folder: options?.folder,
          unreadOnly: options?.unreadOnly,
          subjectContains: options?.subjectContains,
          fromContains: options?.fromContains,
          receivedAfter,
          top: 10,
          orderBy: 'receivedDateTime desc',
        });

        for (const msg of messages) {
          let matched = true;

          if (options?.subjectContains) {
            const hasSubj = msg.subject
              .toLowerCase()
              .includes(options.subjectContains.toLowerCase());
            if (!hasSubj) matched = false;
          }

          if (matched && options?.fromContains) {
            const sender = (msg.from?.emailAddress.address || '').toLowerCase();
            if (!sender.includes(options.fromContains.toLowerCase())) {
              matched = false;
            }
          }

          if (matched && options?.predicate) {
            matched = options.predicate(msg);
          }

          if (matched) {
            if (options?.markAsReadAfterMatch) {
              await this.markAsRead(account, msg.id, true).catch(() => {});
            }
            return msg;
          }
        }
      } catch (err: unknown) {
        if (err instanceof TimeoutError) throw err;
        if (err instanceof TokenRefreshError) throw err;
        if (err instanceof GraphApiError) throw err;
        // RateLimitError (already retried internally) and any other
        // unexpected error: keep polling until timeout.
      }

      const { promise, resolve } = Promise.withResolvers<void>();
      setTimeout(resolve, intervalMs);
      await promise;
    }
  }

  public extractOtp(
    input: OutlookMessage | string,
    options?: OtpExtractOptions
  ): OtpResult | null {
    return extractOtp(input, options);
  }

  private mapToOutlookMessage(raw: Record<string, unknown>): OutlookMessage {
    const fromRaw = raw['from'] as
      | { emailAddress?: { name?: string; address?: string } }
      | undefined;
    const bodyRaw = (raw['body'] as {
      contentType?: 'text' | 'html';
      content?: string;
    }) || { contentType: 'text', content: '' };

    const toRecipientsRaw =
      (raw['toRecipients'] as Array<{
        emailAddress?: { name?: string; address?: string };
      }>) || [];
    const ccRecipientsRaw =
      (raw['ccRecipients'] as Array<{
        emailAddress?: { name?: string; address?: string };
      }>) || [];

    return {
      id: (raw['id'] as string) || '',
      conversationId: raw['conversationId'] as string | undefined,
      subject: (raw['subject'] as string) || '(No Subject)',
      bodyPreview: raw['bodyPreview'] as string | undefined,
      body: {
        contentType: bodyRaw.contentType || 'text',
        content: bodyRaw.content || '',
      },
      from: fromRaw?.emailAddress
        ? {
            emailAddress: {
              name: fromRaw.emailAddress.name,
              address: fromRaw.emailAddress.address || '',
            },
          }
        : undefined,
      toRecipients: toRecipientsRaw
        .filter((r) => r.emailAddress?.address)
        .map((r) => ({
          emailAddress: {
            name: r.emailAddress?.name,
            address: r.emailAddress!.address!,
          },
        })),
      ccRecipients: ccRecipientsRaw
        .filter((r) => r.emailAddress?.address)
        .map((r) => ({
          emailAddress: {
            name: r.emailAddress?.name,
            address: r.emailAddress!.address!,
          },
        })),
      receivedDateTime: (raw['receivedDateTime'] as string) || new Date().toISOString(),
      sentDateTime: raw['sentDateTime'] as string | undefined,
      hasAttachments: Boolean(raw['hasAttachments']),
      isRead: Boolean(raw['isRead']),
      webLink: raw['webLink'] as string | undefined,
      raw,
    };
  }
}
