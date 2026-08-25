import { describe, it, expect, vi, beforeEach } from 'vitest';
import { OutlookTokenManager } from '../../ts/auth/token-manager.js';
import { TokenRefreshError, RateLimitError } from '../../ts/types/errors.js';
import type { AccountCredentials } from '../../ts/types/account.js';

describe('token-manager', () => {
  const mockAccount: AccountCredentials = {
    email: 'test@hotmail.com',
    password: 'password123',
    refreshToken: 'mock-refresh-token-1',
    clientId: 'mock-client-id',
  };

  beforeEach(() => {
    vi.restoreAllMocks();
  });

  it('should successfully refresh access token and update cache', async () => {
    const mockResponse = {
      token_type: 'Bearer',
      scope: 'https://graph.microsoft.com/Mail.Read',
      expires_in: 3600,
      access_token: 'mock-access-token-xyz',
      refresh_token: 'mock-refresh-token-2',
    };

    vi.spyOn(globalThis, 'fetch').mockResolvedValueOnce(
      new Response(JSON.stringify(mockResponse), {
        status: 200,
        headers: { 'Content-Type': 'application/json' },
      })
    );

    const onTokenRotated = vi.fn();
    const tokenManager = new OutlookTokenManager({ onTokenRotated });

    const token = await tokenManager.getAccessToken(mockAccount);

    expect(token).toBe('mock-access-token-xyz');
    expect(mockAccount.refreshToken).toBe('mock-refresh-token-2');
    expect(onTokenRotated).toHaveBeenCalledWith(
      'mock-refresh-token-1',
      'mock-refresh-token-2',
      mockAccount
    );

    const cached = tokenManager.getCachedToken(mockAccount.email);
    expect(cached).toBeDefined();
    expect(cached?.accessToken).toBe('mock-access-token-xyz');
  });

  it('should return cached token if still valid and not forced', async () => {
    const mockResponse = {
      token_type: 'Bearer',
      expires_in: 3600,
      access_token: 'cached-token-123',
    };

    const fetchSpy = vi.spyOn(globalThis, 'fetch').mockResolvedValue(
      new Response(JSON.stringify(mockResponse), {
        status: 200,
        headers: { 'Content-Type': 'application/json' },
      })
    );

    const tokenManager = new OutlookTokenManager();
    const token1 = await tokenManager.getAccessToken(mockAccount);
    const token2 = await tokenManager.getAccessToken(mockAccount);

    expect(token1).toBe('cached-token-123');
    expect(token2).toBe('cached-token-123');
    expect(fetchSpy).toHaveBeenCalledTimes(1);
  });

  it('should refresh token when forceRefresh is true', async () => {
    const fetchSpy = vi
      .spyOn(globalThis, 'fetch')
      .mockResolvedValueOnce(
        new Response(
          JSON.stringify({
            token_type: 'Bearer',
            expires_in: 3600,
            access_token: 'token-1',
          }),
          { status: 200, headers: { 'Content-Type': 'application/json' } }
        )
      )
      .mockResolvedValueOnce(
        new Response(
          JSON.stringify({
            token_type: 'Bearer',
            expires_in: 3600,
            access_token: 'token-2',
          }),
          { status: 200, headers: { 'Content-Type': 'application/json' } }
        )
      );

    const tokenManager = new OutlookTokenManager();
    const token1 = await tokenManager.getAccessToken(mockAccount);
    const token2 = await tokenManager.getAccessToken(mockAccount, true);

    expect(token1).toBe('token-1');
    expect(token2).toBe('token-2');
    expect(fetchSpy).toHaveBeenCalledTimes(2);
  });

  it('should throw TokenRefreshError on invalid_grant / 400 Bad Request', async () => {
    vi.spyOn(globalThis, 'fetch').mockResolvedValueOnce(
      new Response(
        JSON.stringify({
          error: 'invalid_grant',
          error_description: 'The provided refresh token is expired or revoked.',
        }),
        { status: 400, headers: { 'Content-Type': 'application/json' } }
      )
    );

    const tokenManager = new OutlookTokenManager();
    await expect(tokenManager.refreshAccessToken(mockAccount)).rejects.toThrow(
      TokenRefreshError
    );
  });

  it('should retry on rate limit 429', async () => {
    const fetchSpy = vi
      .spyOn(globalThis, 'fetch')
      .mockResolvedValueOnce(
        new Response('Too Many Requests', {
          status: 429,
          headers: { 'Retry-After': '1' },
        })
      )
      .mockResolvedValueOnce(
        new Response(
          JSON.stringify({
            token_type: 'Bearer',
            expires_in: 3600,
            access_token: 'recovered-token',
          }),
          { status: 200, headers: { 'Content-Type': 'application/json' } }
        )
      );

    const tokenManager = new OutlookTokenManager({ initialRetryDelayMs: 10 });
    const response = await tokenManager.refreshAccessToken(mockAccount);

    expect(response.access_token).toBe('recovered-token');
    expect(fetchSpy).toHaveBeenCalledTimes(2);
  });
});
