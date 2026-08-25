import type { AccountCredentials } from './account.js';

export interface TokenResponse {
  token_type: string;
  scope: string;
  expires_in: number;
  ext_expires_in?: number;
  access_token: string;
  refresh_token?: string;
  id_token?: string;
}

export interface TokenCacheEntry {
  accessToken: string;
  refreshToken: string;
  expiresAt: number;
  scope: string;
}

export type TokenRotationCallback = (
  oldRefreshToken: string,
  newRefreshToken: string,
  account: AccountCredentials
) => void | Promise<void>;

export interface TokenManagerOptions {
  authority?: string;
  scopes?: string[];
  expiryBufferSec?: number;
  maxRetries?: number;
  initialRetryDelayMs?: number;
  requestTimeoutMs?: number;
  onTokenRotated?: TokenRotationCallback;
}
