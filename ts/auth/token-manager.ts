import { TokenRefreshError, RateLimitError } from '../types/errors.js';
import { parseOAuthError } from './oauth-error.js';
import type { AccountCredentials } from '../types/account.js';
import type {
  TokenResponse,
  TokenCacheEntry,
  TokenManagerOptions,
} from '../types/auth.js';

export class OutlookTokenManager {
  private readonly tokenCache = new Map<string, TokenCacheEntry>();
  private readonly inFlightRefreshes = new Map<string, Promise<TokenResponse>>();
  private readonly defaultAuthority: string;
  private readonly defaultScopes: string[];
  private readonly expiryBufferMs: number;
  private readonly maxRetries: number;
  private readonly initialRetryDelayMs: number;
  private readonly requestTimeoutMs: number;
  private readonly onTokenRotated?: TokenManagerOptions['onTokenRotated'];

  constructor(options?: TokenManagerOptions) {
    this.defaultAuthority = options?.authority ?? 'common';
    this.defaultScopes = options?.scopes ?? [
      'https://graph.microsoft.com/Mail.Read',
      'https://graph.microsoft.com/Mail.ReadWrite',
      'offline_access',
    ];
    this.expiryBufferMs = (options?.expiryBufferSec ?? 300) * 1000;
    this.maxRetries = options?.maxRetries ?? 3;
    this.initialRetryDelayMs = options?.initialRetryDelayMs ?? 1000;
    this.requestTimeoutMs = options?.requestTimeoutMs ?? 15000;
    this.onTokenRotated = options?.onTokenRotated;
  }

  public getCachedToken(email: string): TokenCacheEntry | undefined {
    return this.tokenCache.get(email.toLowerCase());
  }

  public clearCache(email?: string): void {
    if (email) {
      this.tokenCache.delete(email.toLowerCase());
    } else {
      this.tokenCache.clear();
    }
  }

  public isTokenValid(entry: TokenCacheEntry): boolean {
    return Date.now() < entry.expiresAt - this.expiryBufferMs;
  }

  public async getAccessToken(
    account: AccountCredentials,
    forceRefresh = false
  ): Promise<string> {
    const key = account.email.toLowerCase();
    const cached = this.tokenCache.get(key);

    if (!forceRefresh && cached && this.isTokenValid(cached)) {
      return cached.accessToken;
    }

    const response = await this.refreshAccessToken(account);
    return response.access_token;
  }

  public async refreshAccessToken(
    account: AccountCredentials
  ): Promise<TokenResponse> {
    const key = account.email.toLowerCase();
    const existingInFlight = this.inFlightRefreshes.get(key);
    if (existingInFlight) {
      return existingInFlight;
    }

    const refreshPromise = this.executeRefresh(account);
    this.inFlightRefreshes.set(key, refreshPromise);

    try {
      return await refreshPromise;
    } finally {
      this.inFlightRefreshes.delete(key);
    }
  }

  private async executeRefresh(
    account: AccountCredentials
  ): Promise<TokenResponse> {
    const authority = account.authority ?? this.defaultAuthority;
    const tokenUrl = `https://login.microsoftonline.com/${authority}/oauth2/v2.0/token`;
    const scope = this.defaultScopes.join(' ');

    const bodyParams = new URLSearchParams({
      client_id: account.clientId,
      grant_type: 'refresh_token',
      refresh_token: account.refreshToken,
      scope,
    });

    let attempt = 0;
    let delay = this.initialRetryDelayMs;

    while (attempt <= this.maxRetries) {
      attempt++;
      try {
        const res = await fetch(tokenUrl, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/x-www-form-urlencoded',
            Accept: 'application/json',
          },
          body: bodyParams.toString(),
          signal: AbortSignal.timeout(this.requestTimeoutMs),
        });

        if (res.status === 429) {
          const retryAfterHeader = res.headers.get('Retry-After');
          const retryAfterSec = retryAfterHeader ? parseInt(retryAfterHeader, 10) : NaN;
          const retryDelayMs = !isNaN(retryAfterSec) && retryAfterSec > 0
            ? retryAfterSec * 1000
            : delay;

          if (attempt > this.maxRetries) {
            throw new RateLimitError(
              `Rate limited (HTTP 429) refreshing token for account ${account.email}. Max retries exceeded.`,
              retryDelayMs,
              429
            );
          }

          const { promise, resolve } = Promise.withResolvers<void>();
          setTimeout(resolve, retryDelayMs + Math.random() * 200);
          await promise;

          delay *= 2;
          continue;
        }

        if (res.status >= 500 && res.status < 600) {
          if (attempt > this.maxRetries) {
            const errorText = await res.text();
            throw new TokenRefreshError(
              `Server error (HTTP ${res.status}) while refreshing token for account ${account.email}: ${errorText}`,
              res.status
            );
          }

          const { promise, resolve } = Promise.withResolvers<void>();
          setTimeout(resolve, delay + Math.random() * 200);
          await promise;

          delay *= 2;
          continue;
        }

        const data = (await res.json()) as Record<string, unknown>;

        if (!res.ok) {
          throw parseOAuthError(res.status, data, account.email);
        }

        const tokenResponse: TokenResponse = {
          token_type: (data['token_type'] as string) || 'Bearer',
          scope: (data['scope'] as string) || scope,
          expires_in: (data['expires_in'] as number) || 3600,
          ext_expires_in: data['ext_expires_in'] as number | undefined,
          access_token: data['access_token'] as string,
          refresh_token: data['refresh_token'] as string | undefined,
          id_token: data['id_token'] as string | undefined,
        };

        const oldRefreshToken = account.refreshToken;
        if (tokenResponse.refresh_token && tokenResponse.refresh_token !== oldRefreshToken) {
          account.refreshToken = tokenResponse.refresh_token;
          if (this.onTokenRotated) {
            try {
              await this.onTokenRotated(
                oldRefreshToken,
                tokenResponse.refresh_token,
                account
              );
            } catch (listenerErr) {
              console.warn(
                `Warning: onTokenRotated listener failed for ${account.email}:`,
                listenerErr
              );
            }
          }
        }

        const expiresAt = Date.now() + tokenResponse.expires_in * 1000;
        this.tokenCache.set(account.email.toLowerCase(), {
          accessToken: tokenResponse.access_token,
          refreshToken: tokenResponse.refresh_token ?? account.refreshToken,
          expiresAt,
          scope: tokenResponse.scope,
        });

        return tokenResponse;
      } catch (err: unknown) {
        if (
          err instanceof TokenRefreshError ||
          err instanceof RateLimitError
        ) {
          throw err;
        }

        if (attempt > this.maxRetries) {
          const message = err instanceof Error ? err.message : String(err);
          throw new TokenRefreshError(
            `Network error refreshing token for ${account.email} after ${attempt} attempts: ${message}`
          );
        }

        const { promise, resolve } = Promise.withResolvers<void>();
        setTimeout(resolve, delay + Math.random() * 200);
        await promise;

        delay *= 2;
      }
    }

    throw new TokenRefreshError(
      `Failed to refresh token for ${account.email}: max retries exceeded`
    );
  }
}
