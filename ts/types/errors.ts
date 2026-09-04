export class GraphModuleError extends Error {
  constructor(message: string, public readonly code?: string) {
    super(message);
    this.name = 'GraphModuleError';
  }
}

export class AccountParseError extends GraphModuleError {
  constructor(
    message: string,
    public readonly rawLine?: string,
    public readonly lineNumber?: number
  ) {
    super(message, 'ACCOUNT_PARSE_ERROR');
    this.name = 'AccountParseError';
  }
}

export type TokenErrorReason =
  | 'EXPIRED_OR_REVOKED_GRANT'
  | 'CLIENT_ID_MISMATCH'
  | 'INVALID_CLIENT'
  | 'ACCOUNT_LOCKED'
  | 'ACCOUNT_DISABLED'
  | 'INTERACTION_REQUIRED'
  | 'INVALID_SCOPE'
  | 'UNKNOWN';

export interface TokenErrorDiagnostics {
  aadstsCode?: number;
  diagnosticReason?: TokenErrorReason;
  remediation?: string;
}

export class TokenRefreshError extends GraphModuleError {
  public readonly aadstsCode?: number;
  public readonly diagnosticReason?: TokenErrorReason;
  public readonly remediation?: string;

  constructor(
    message: string,
    public readonly statusCode?: number,
    public readonly errorResponse?: Record<string, unknown>,
    diagnostics?: TokenErrorDiagnostics
  ) {
    super(message, 'TOKEN_REFRESH_ERROR');
    this.name = 'TokenRefreshError';
    this.aadstsCode = diagnostics?.aadstsCode;
    this.diagnosticReason = diagnostics?.diagnosticReason;
    this.remediation = diagnostics?.remediation;
  }
}

export class GraphApiError extends GraphModuleError {
  constructor(
    message: string,
    public readonly statusCode?: number,
    public readonly graphErrorCode?: string,
    public readonly responseBody?: unknown
  ) {
    super(message, 'GRAPH_API_ERROR');
    this.name = 'GraphApiError';
  }
}

export class RateLimitError extends GraphModuleError {
  constructor(
    message: string,
    public readonly retryAfterMs: number,
    public readonly statusCode: number = 429
  ) {
    super(message, 'RATE_LIMIT_ERROR');
    this.name = 'RateLimitError';
  }
}

export class TimeoutError extends GraphModuleError {
  constructor(message: string) {
    super(message, 'TIMEOUT_ERROR');
    this.name = 'TimeoutError';
  }
}
