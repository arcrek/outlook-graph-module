export interface OtpExtractOptions {
  preferredLength?: number;
  allowAlphanumeric?: boolean;
  customPatterns?: RegExp[];
  locale?: 'en' | 'vi' | 'auto';
}

export interface OtpResult {
  code: string;
  digits: number;
  confidence: number;
  patternMatched: string;
  contextSnippet?: string;
  rawText: string;
}
