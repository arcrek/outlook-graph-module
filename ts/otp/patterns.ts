export const KEYWORDS_EN = [
  'verification code',
  'security code',
  'confirmation code',
  'one-time password',
  'one-time passcode',
  'otp',
  'passcode',
  'secret code',
  'access code',
  'login code',
  'authorization code',
  'auth code',
  'validation code',
  'pin code',
  'temporary password',
];

export const KEYWORDS_VI = [
  'mã xác thực',
  'mã xác nhận',
  'mã otp',
  'mã bảo mật',
  'mã kiểm tra',
  'mã xác minh',
  'mật khẩu một lần',
  'mã kích hoạt',
  'mã an toàn',
  'mã giao dịch',
  'mã đăng nhập',
  'mã số bí mật',
];

export const CONTEXT_PREFIX_REGEX = new RegExp(
  `(?:(?:${[...KEYWORDS_EN, ...KEYWORDS_VI].join('|')})[^0-9a-zA-Z]{0,40}?(?:is|là|:|=|-|\\s)?\\s*)([0-9]{4,8}|[0-9]{3}[-\\s][0-9]{3}|[A-Z0-9]{6,8})\\b`,
  'i'
);

export const CONTEXT_SUFFIX_REGEX = new RegExp(
  `\\b([0-9]{4,8}|[0-9]{3}[-\\s][0-9]{3}|[A-Z0-9]{6,8})\\s*(?:is\\s*(?:your|the)?\\s*(?:verification|security|confirmation|otp)?\\s*code|là\\s*mã\\s*xác\\s*(?:thực|nhận|minh))\\b`,
  'i'
);

export const NUMERIC_CODE_REGEXES = [
  { pattern: /\b([0-9]{6})\b/g, digits: 6, weight: 0.9 },
  { pattern: /\b([0-9]{3}[-\s][0-9]{3})\b/g, digits: 6, weight: 0.95 },
  { pattern: /\b([0-9]{4})\b/g, digits: 4, weight: 0.7 },
  { pattern: /\b([0-9]{8})\b/g, digits: 8, weight: 0.8 },
  { pattern: /\b([0-9]{5})\b/g, digits: 5, weight: 0.65 },
  { pattern: /\b([0-9]{7})\b/g, digits: 7, weight: 0.65 },
];

export const ALPHANUMERIC_CODE_REGEX = /\b([A-Z0-9]{6,8})\b/g;

export const FALSE_POSITIVE_YEAR_REGEX = /\b(19\d\d|20\d\d)\b/;
export const COPYRIGHT_YEAR_REGEX = /(?:copyright|©|\(c\))\s*(?:19\d\d|20\d\d)/i;
