// Core Classes
export { OutlookMailClient } from './mail/mail-client.js';
export { OutlookTokenManager } from './auth/token-manager.js';
export { GraphHttpClient } from './http/graph-client.js';

// Parser Functions
export {
  parseAccountLine,
  parseAccountLines,
  parseAccountFile,
} from './parser/account-parser.js';

// OTP Engine
export {
  extractOtp,
  extractAllOtps,
  cleanHtml,
} from './otp/otp-extractor.js';

// Query Builder
export {
  buildODataQuery,
  escapeODataString,
  getMailEndpoint,
} from './http/query-builder.js';

// Error Classes
export {
  GraphModuleError,
  AccountParseError,
  TokenRefreshError,
  GraphApiError,
  RateLimitError,
  TimeoutError,
} from './types/errors.js';

// Types & Interfaces
export type {
  AccountCredentials,
  ParseAccountOptions,
  TokenResponse,
  TokenCacheEntry,
  TokenManagerOptions,
  TokenRotationCallback,
  OutlookMessage,
  OutlookAttachment,
  EmailAddress,
  EmailRecipient,
  ItemBody,
  GetMessagesOptions,
  WaitForEmailOptions,
  OtpExtractOptions,
  OtpResult,
} from './types/index.js';
