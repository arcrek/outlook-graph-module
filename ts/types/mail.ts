export interface EmailAddress {
  name?: string;
  address: string;
}

export interface EmailRecipient {
  emailAddress: EmailAddress;
}

export interface ItemBody {
  contentType: 'text' | 'html';
  content: string;
}

export interface OutlookAttachment {
  id: string;
  name: string;
  contentType: string;
  size: number;
  isInline: boolean;
  contentBytes?: string;
  data?: Buffer;
}

export interface OutlookMessage {
  id: string;
  conversationId?: string;
  subject: string;
  bodyPreview?: string;
  body: ItemBody;
  from?: EmailRecipient;
  toRecipients?: EmailRecipient[];
  ccRecipients?: EmailRecipient[];
  receivedDateTime: string;
  sentDateTime?: string;
  hasAttachments: boolean;
  isRead: boolean;
  webLink?: string;
  attachments?: OutlookAttachment[];
  raw?: Record<string, unknown>;
}

export interface GetMessagesOptions {
  folder?: string;
  top?: number;
  skip?: number;
  unreadOnly?: boolean;
  subjectContains?: string;
  fromContains?: string;
  receivedAfter?: Date | string;
  search?: string;
  orderBy?: string;
  select?: string[];
  filter?: string;
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
