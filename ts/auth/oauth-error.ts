import {
  TokenRefreshError,
  type TokenErrorDiagnostics,
  type TokenErrorReason,
} from '../types/errors.js';

export function parseOAuthError(
  statusCode: number,
  data: Record<string, unknown> | undefined,
  email: string
): TokenRefreshError {
  const errorDesc =
    (data?.['error_description'] as string) ||
    (data?.['error'] as string) ||
    'Bad Request';

  let aadstsCode: number | undefined;

  // Extract from error_codes array if provided by Microsoft
  if (Array.isArray(data?.['error_codes']) && data['error_codes'].length > 0) {
    const rawCode = data['error_codes'][0];
    const parsed = typeof rawCode === 'number' ? rawCode : parseInt(String(rawCode), 10);
    if (!isNaN(parsed)) {
      aadstsCode = parsed;
    }
  }

  // Fallback: extract AADSTS code from error description via regex
  if (aadstsCode === undefined) {
    const match = errorDesc.match(/AADSTS(\d+)/i);
    if (match && match[1]) {
      aadstsCode = parseInt(match[1], 10);
    }
  }

  let diagnosticReason: TokenErrorReason = 'UNKNOWN';
  let remediation =
    'Token refresh failed. Verify account credentials and Azure AD configuration.';

  if (aadstsCode === 70000) {
    if (/different client/i.test(errorDesc)) {
      diagnosticReason = 'CLIENT_ID_MISMATCH';
      remediation =
        'The refresh token was issued for a different client_id. Verify that client_id matches the application that generated the token.';
    } else {
      diagnosticReason = 'EXPIRED_OR_REVOKED_GRANT';
      remediation =
        'The refresh token is expired or revoked. Re-authenticate to obtain a new refresh token.';
    }
  } else if (
    aadstsCode === 700082 ||
    aadstsCode === 70008 ||
    aadstsCode === 700084 ||
    aadstsCode === 50173
  ) {
    diagnosticReason = 'EXPIRED_OR_REVOKED_GRANT';
    remediation =
      'The refresh token has expired or was revoked. Re-authenticate to obtain a new refresh token.';
  } else if (aadstsCode === 700016) {
    diagnosticReason = 'INVALID_CLIENT';
    remediation =
      'The client_id was not found in Azure AD. Verify your client_id configuration.';
  } else if (aadstsCode === 50053) {
    diagnosticReason = 'ACCOUNT_LOCKED';
    remediation =
      'The account is locked due to multiple failed sign-in attempts. Unlock it via Microsoft account security.';
  } else if (aadstsCode === 50057) {
    diagnosticReason = 'ACCOUNT_DISABLED';
    remediation =
      'The account is disabled. Contact Microsoft support or your administrator.';
  } else if (
    aadstsCode === 50076 ||
    aadstsCode === 50079 ||
    aadstsCode === 53003
  ) {
    diagnosticReason = 'INTERACTION_REQUIRED';
    remediation =
      'Interactive sign-in or multi-factor authentication (MFA) is required by security policy.';
  } else if (aadstsCode === 70011) {
    diagnosticReason = 'INVALID_SCOPE';
    remediation =
      'The requested OAuth scopes are invalid for this resource. Check requested OAuth scopes.';
  } else if (data?.['error'] === 'invalid_grant') {
    diagnosticReason = 'EXPIRED_OR_REVOKED_GRANT';
    remediation =
      'The grant or refresh token is invalid, expired, or revoked. Re-authenticate to obtain a new token.';
  }

  const diagnostics: TokenErrorDiagnostics = {
    aadstsCode,
    diagnosticReason,
    remediation,
  };

  const message = `Failed to refresh OAuth token for account ${email}: ${errorDesc}`;
  return new TokenRefreshError(message, statusCode, data, diagnostics);
}
