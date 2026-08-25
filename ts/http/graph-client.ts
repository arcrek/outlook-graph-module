import { GraphApiError, RateLimitError } from '../types/errors.js';
import type { AccountCredentials } from '../types/account.js';
import type { OutlookTokenManager } from '../auth/token-manager.js';

export interface RequestOptions {
  headers?: Record<string, string>;
  body?: unknown;
  searchConsistencyEventual?: boolean;
  timeoutMs?: number;
  maxRetries?: number;
  initialRetryDelayMs?: number;
}

export class GraphHttpClient {
  private readonly defaultTimeoutMs: number;
  private readonly defaultMaxRetries: number;
  private readonly defaultInitialRetryDelayMs: number;

  constructor(
    private readonly tokenManager: OutlookTokenManager,
    options?: {
      timeoutMs?: number;
      maxRetries?: number;
      initialRetryDelayMs?: number;
    }
  ) {
    this.defaultTimeoutMs = options?.timeoutMs ?? 20000;
    this.defaultMaxRetries = options?.maxRetries ?? 3;
    this.defaultInitialRetryDelayMs = options?.initialRetryDelayMs ?? 1000;
  }

  public async request<T>(
    account: AccountCredentials,
    url: string,
    method: 'GET' | 'POST' | 'PATCH' | 'DELETE' = 'GET',
    options?: RequestOptions
  ): Promise<T> {
    let accessToken = await this.tokenManager.getAccessToken(account);
    let didRetryAuth = false;
    let attempt = 0;
    const maxRetries = options?.maxRetries ?? this.defaultMaxRetries;
    let retryDelay = options?.initialRetryDelayMs ?? this.defaultInitialRetryDelayMs;
    const timeoutMs = options?.timeoutMs ?? this.defaultTimeoutMs;

    while (attempt <= maxRetries) {
      attempt++;
      const headers: Record<string, string> = {
        Authorization: `Bearer ${accessToken}`,
        Accept: 'application/json',
        ...options?.headers,
      };

      if (options?.searchConsistencyEventual) {
        headers['ConsistencyLevel'] = 'eventual';
      }

      let requestBody: string | undefined;
      if (options?.body !== undefined) {
        headers['Content-Type'] = 'application/json';
        requestBody = JSON.stringify(options.body);
      }

      try {
        const res = await fetch(url, {
          method,
          headers,
          body: requestBody,
          signal: AbortSignal.timeout(timeoutMs),
        });

        if (res.status === 401 && !didRetryAuth) {
          didRetryAuth = true;
          accessToken = await this.tokenManager.getAccessToken(account, true);
          continue;
        }

        if (res.status === 429) {
          const retryAfterHeader = res.headers.get('Retry-After');
          const retryAfterSec = retryAfterHeader ? parseInt(retryAfterHeader, 10) : NaN;
          const retryDelayMs =
            !isNaN(retryAfterSec) && retryAfterSec > 0
              ? retryAfterSec * 1000
              : retryDelay;

          if (attempt > maxRetries) {
            throw new RateLimitError(
              `Microsoft Graph rate limit reached for account ${account.email}. Max retries exceeded.`,
              retryDelayMs,
              429
            );
          }

          const { promise, resolve } = Promise.withResolvers<void>();
          setTimeout(resolve, retryDelayMs + Math.random() * 200);
          await promise;

          retryDelay *= 2;
          continue;
        }

        if (res.status >= 500 && res.status < 600) {
          if (attempt > maxRetries) {
            const errorText = await res.text();
            throw new GraphApiError(
              `Microsoft Graph server error (${res.status}): ${errorText}`,
              res.status,
              'ServerError'
            );
          }

          const { promise, resolve } = Promise.withResolvers<void>();
          setTimeout(resolve, retryDelay + Math.random() * 200);
          await promise;

          retryDelay *= 2;
          continue;
        }

        if (res.status === 204) {
          return undefined as unknown as T;
        }

        const contentType = res.headers.get('content-type') || '';
        const isJson = contentType.includes('application/json');

        if (!res.ok) {
          let errorCode = 'UnknownError';
          let errorMessage = res.statusText;
          let responseBody: unknown;

          if (isJson) {
            try {
              const data = (await res.json()) as Record<string, unknown>;
              responseBody = data;
              const errObj = data['error'] as Record<string, unknown> | undefined;
              if (errObj) {
                errorCode = (errObj['code'] as string) || errorCode;
                errorMessage = (errObj['message'] as string) || errorMessage;
              }
            } catch {
              // Ignore JSON parse failure
            }
          } else {
            errorMessage = await res.text();
          }

          throw new GraphApiError(
            `Microsoft Graph API error (${res.status} ${errorCode}): ${errorMessage}`,
            res.status,
            errorCode,
            responseBody
          );
        }

        if (isJson) {
          return (await res.json()) as T;
        }

        return (await res.text()) as unknown as T;
      } catch (err: unknown) {
        if (
          err instanceof GraphApiError ||
          err instanceof RateLimitError
        ) {
          throw err;
        }

        if (attempt > maxRetries) {
          const message = err instanceof Error ? err.message : String(err);
          throw new GraphApiError(
            `Network request failed to '${url}' after ${attempt} attempts: ${message}`
          );
        }

        const { promise, resolve } = Promise.withResolvers<void>();
        setTimeout(resolve, retryDelay + Math.random() * 200);
        await promise;

        retryDelay *= 2;
      }
    }

    throw new GraphApiError(`Request to '${url}' failed: max retries exceeded`);
  }

  public async get<T>(
    account: AccountCredentials,
    url: string,
    options?: RequestOptions
  ): Promise<T> {
    return this.request<T>(account, url, 'GET', options);
  }

  public async post<T>(
    account: AccountCredentials,
    url: string,
    body?: unknown,
    options?: Omit<RequestOptions, 'body'>
  ): Promise<T> {
    return this.request<T>(account, url, 'POST', { ...options, body });
  }

  public async patch<T>(
    account: AccountCredentials,
    url: string,
    body?: unknown,
    options?: Omit<RequestOptions, 'body'>
  ): Promise<T> {
    return this.request<T>(account, url, 'PATCH', { ...options, body });
  }

  public async delete<T>(
    account: AccountCredentials,
    url: string,
    options?: RequestOptions
  ): Promise<T> {
    return this.request<T>(account, url, 'DELETE', options);
  }
}
