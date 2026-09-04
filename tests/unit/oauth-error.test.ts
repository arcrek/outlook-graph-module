import { describe, it, expect } from 'vitest';
import { parseOAuthError } from '../../ts/auth/oauth-error.js';
import { TokenRefreshError } from '../../ts/types/errors.js';

describe('oauth-error diagnostics', () => {
  it('should classify AADSTS70000 expired grant as EXPIRED_OR_REVOKED_GRANT', () => {
    const err = parseOAuthError(
      400,
      {
        error: 'invalid_grant',
        error_description:
          'AADSTS70000: The user could not be authenticated as the grant is expired. The user must sign in again.',
        error_codes: [70000],
      },
      'test@hotmail.com'
    );

    expect(err).toBeInstanceOf(TokenRefreshError);
    expect(err.statusCode).toBe(400);
    expect(err.aadstsCode).toBe(70000);
    expect(err.diagnosticReason).toBe('EXPIRED_OR_REVOKED_GRANT');
    expect(err.remediation).toContain('expired or revoked');
    expect(err.message).toContain('test@hotmail.com');
  });

  it('should classify AADSTS70000 different client as CLIENT_ID_MISMATCH', () => {
    const err = parseOAuthError(
      400,
      {
        error: 'invalid_grant',
        error_description:
          'AADSTS70000: The token was issued for a different client id',
      },
      'test@hotmail.com'
    );

    expect(err.aadstsCode).toBe(70000);
    expect(err.diagnosticReason).toBe('CLIENT_ID_MISMATCH');
    expect(err.remediation).toContain('different client_id');
  });

  it('should classify AADSTS700082 as EXPIRED_OR_REVOKED_GRANT', () => {
    const err = parseOAuthError(
      400,
      {
        error: 'invalid_grant',
        error_description: 'AADSTS700082: The refresh token has expired',
      },
      'test@hotmail.com'
    );

    expect(err.aadstsCode).toBe(700082);
    expect(err.diagnosticReason).toBe('EXPIRED_OR_REVOKED_GRANT');
  });

  it('should classify AADSTS700016 as INVALID_CLIENT', () => {
    const err = parseOAuthError(
      400,
      {
        error: 'invalid_client',
        error_description:
          'AADSTS700016: Application with identifier was not found in directory',
        error_codes: [700016],
      },
      'test@hotmail.com'
    );

    expect(err.aadstsCode).toBe(700016);
    expect(err.diagnosticReason).toBe('INVALID_CLIENT');
  });

  it('should classify AADSTS50053 as ACCOUNT_LOCKED', () => {
    const err = parseOAuthError(
      400,
      {
        error: 'invalid_grant',
        error_description: 'AADSTS50053: The account is locked',
      },
      'test@hotmail.com'
    );

    expect(err.aadstsCode).toBe(50053);
    expect(err.diagnosticReason).toBe('ACCOUNT_LOCKED');
    expect(err.remediation).toContain('locked');
  });

  it('should classify AADSTS50057 as ACCOUNT_DISABLED', () => {
    const err = parseOAuthError(
      400,
      {
        error: 'invalid_grant',
        error_description: 'AADSTS50057: The user account is disabled',
      },
      'test@hotmail.com'
    );

    expect(err.aadstsCode).toBe(50057);
    expect(err.diagnosticReason).toBe('ACCOUNT_DISABLED');
  });

  it('should classify AADSTS50076 / MFA as INTERACTION_REQUIRED', () => {
    const err = parseOAuthError(
      400,
      {
        error: 'invalid_grant',
        error_description: 'AADSTS50076: Due to a configuration change MFA is required',
      },
      'test@hotmail.com'
    );

    expect(err.aadstsCode).toBe(50076);
    expect(err.diagnosticReason).toBe('INTERACTION_REQUIRED');
    expect(err.remediation).toContain('Interactive sign-in or multi-factor');
  });

  it('should classify AADSTS70011 as INVALID_SCOPE', () => {
    const err = parseOAuthError(
      400,
      {
        error: 'invalid_scope',
        error_description: 'AADSTS70011: The scope is not valid',
      },
      'test@hotmail.com'
    );

    expect(err.aadstsCode).toBe(70011);
    expect(err.diagnosticReason).toBe('INVALID_SCOPE');
  });

  it('should fallback to EXPIRED_OR_REVOKED_GRANT when error is invalid_grant without AADSTS code', () => {
    const err = parseOAuthError(
      400,
      {
        error: 'invalid_grant',
        error_description: 'The provided refresh token is expired or revoked.',
      },
      'test@hotmail.com'
    );

    expect(err.aadstsCode).toBeUndefined();
    expect(err.diagnosticReason).toBe('EXPIRED_OR_REVOKED_GRANT');
  });

  it('should mark unknown error as UNKNOWN', () => {
    const err = parseOAuthError(
      400,
      {
        error: 'unknown_error',
        error_description: 'Something unexpected happened',
      },
      'test@hotmail.com'
    );

    expect(err.aadstsCode).toBeUndefined();
    expect(err.diagnosticReason).toBe('UNKNOWN');
  });
});
