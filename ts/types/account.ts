export interface AccountCredentials {
  email: string;
  password?: string;
  refreshToken: string;
  clientId: string;
  authority?: string;
}

export interface ParseAccountOptions {
  trim?: boolean;
  ignoreEmptyLines?: boolean;
  ignoreComments?: boolean;
  defaultAuthority?: string;
}
