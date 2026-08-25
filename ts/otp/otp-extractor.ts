import type { OutlookMessage } from '../types/mail.js';
import type { OtpExtractOptions, OtpResult } from '../types/otp.js';
import {
  KEYWORDS_EN,
  KEYWORDS_VI,
  CONTEXT_PREFIX_REGEX,
  CONTEXT_SUFFIX_REGEX,
  NUMERIC_CODE_REGEXES,
  COPYRIGHT_YEAR_REGEX,
} from './patterns.js';

export function cleanHtml(html: string): string {
  if (!html) {
    return '';
  }

  let text = html
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<svg[\s\S]*?<\/svg>/gi, ' ')
    .replace(/<[^>]+>/g, ' ');

  // Decode common HTML entities
  const entities: Record<string, string> = {
    '&nbsp;': ' ',
    '&amp;': '&',
    '&lt;': '<',
    '&gt;': '>',
    '&quot;': '"',
    '&#39;': "'",
    '&apos;': "'",
    '&zwnj;': '',
    '&zwj;': '',
  };

  for (const [entity, replacement] of Object.entries(entities)) {
    text = text.replaceAll(entity, replacement);
  }

  // Numeric decimal & hex entities
  text = text.replace(/&#(\d+);/g, (_, code) =>
    String.fromCharCode(parseInt(code, 10))
  );
  text = text.replace(/&#x([0-9a-fA-F]+);/g, (_, code) =>
    String.fromCharCode(parseInt(code, 16))
  );

  return text.replace(/\s+/g, ' ').trim();
}

function resolveRawText(input: OutlookMessage | string): string {
  if (typeof input === 'string') {
    return input.includes('<') && input.includes('>') ? cleanHtml(input) : input;
  }

  const subject = input.subject || '';
  const bodyContent = input.body?.content || input.bodyPreview || '';
  const cleanedBody =
    input.body?.contentType === 'html' || (bodyContent.includes('<') && bodyContent.includes('>'))
      ? cleanHtml(bodyContent)
      : bodyContent;

  return `${subject}\n${cleanedBody}`;
}

export function extractOtp(
  input: OutlookMessage | string,
  options?: OtpExtractOptions
): OtpResult | null {
  const results = extractAllOtps(input, options);
  return results.length > 0 ? results[0] : null;
}

export function extractAllOtps(
  input: OutlookMessage | string,
  options?: OtpExtractOptions
): OtpResult[] {
  const rawText = resolveRawText(input);
  if (!rawText.trim()) {
    return [];
  }

  const candidates: OtpResult[] = [];
  const seenCodes = new Set<string>();

  // 1. Check direct context prefix regex
  const prefixMatch = rawText.match(CONTEXT_PREFIX_REGEX);
  if (prefixMatch && prefixMatch[1]) {
    const rawCode = prefixMatch[1];
    const normalizedCode = rawCode.replace(/[-\s]/g, '');
    const index = prefixMatch.index ?? 0;
    const snippetStart = Math.max(0, index - 30);
    const snippetEnd = Math.min(rawText.length, index + prefixMatch[0].length + 30);
    const snippet = rawText.slice(snippetStart, snippetEnd).trim();

    seenCodes.add(normalizedCode);
    candidates.push({
      code: normalizedCode,
      digits: normalizedCode.length,
      confidence: 0.98,
      patternMatched: 'context_prefix',
      contextSnippet: snippet,
      rawText: rawCode,
    });
  }

  // 2. Check context suffix regex
  const suffixMatch = rawText.match(CONTEXT_SUFFIX_REGEX);
  if (suffixMatch && suffixMatch[1]) {
    const rawCode = suffixMatch[1];
    const normalizedCode = rawCode.replace(/[-\s]/g, '');
    if (!seenCodes.has(normalizedCode)) {
      const index = suffixMatch.index ?? 0;
      const snippetStart = Math.max(0, index - 30);
      const snippetEnd = Math.min(rawText.length, index + suffixMatch[0].length + 30);
      const snippet = rawText.slice(snippetStart, snippetEnd).trim();

      seenCodes.add(normalizedCode);
      candidates.push({
        code: normalizedCode,
        digits: normalizedCode.length,
        confidence: 0.95,
        patternMatched: 'context_suffix',
        contextSnippet: snippet,
        rawText: rawCode,
      });
    }
  }

  // 3. Keyword proximity search with general numeric patterns
  const allKeywords =
    options?.locale === 'en'
      ? KEYWORDS_EN
      : options?.locale === 'vi'
      ? KEYWORDS_VI
      : [...KEYWORDS_EN, ...KEYWORDS_VI];

  const lowerText = rawText.toLowerCase();

  for (const { pattern, digits, weight } of NUMERIC_CODE_REGEXES) {
    if (options?.preferredLength && options.preferredLength !== digits) {
      continue;
    }

    const matches = Array.from(rawText.matchAll(pattern));
    for (const match of matches) {
      const rawCode = match[1];
      const normalizedCode = rawCode.replace(/[-\s]/g, '');
      const matchIndex = match.index ?? 0;

      // Filter out years (1900-2099) if part of copyright or dates without OTP context
      const isYear = (digits === 4 && (normalizedCode.startsWith('19') || normalizedCode.startsWith('20')));
      if (isYear) {
        const surroundingText = rawText.slice(
          Math.max(0, matchIndex - 40),
          Math.min(rawText.length, matchIndex + 40)
        );
        if (COPYRIGHT_YEAR_REGEX.test(surroundingText)) {
          continue;
        }
      }

      // Calculate proximity to any keyword
      let minDistance = Infinity;
      let closestKeyword = '';

      for (const kw of allKeywords) {
        let kwPos = lowerText.indexOf(kw);
        while (kwPos !== -1) {
          const dist = Math.abs(matchIndex - kwPos);
          if (dist < minDistance) {
            minDistance = dist;
            closestKeyword = kw;
          }
          kwPos = lowerText.indexOf(kw, kwPos + 1);
        }
      }

      let confidence = weight;
      if (minDistance < 60) {
        confidence = Math.min(0.99, weight + 0.15);
      } else if (minDistance < 150) {
        confidence = Math.min(0.85, weight + 0.05);
      } else if (minDistance < 300) {
        confidence = weight * 0.7;
      } else {
        confidence = isYear ? 0.1 : weight * 0.4;
      }

      if (seenCodes.has(normalizedCode)) {
        // Upgrade confidence if already seen
        const existing = candidates.find((c) => c.code === normalizedCode);
        if (existing && confidence > existing.confidence) {
          existing.confidence = confidence;
        }
        continue;
      }

      seenCodes.add(normalizedCode);
      const snippetStart = Math.max(0, matchIndex - 30);
      const snippetEnd = Math.min(rawText.length, matchIndex + rawCode.length + 30);
      const snippet = rawText.slice(snippetStart, snippetEnd).trim();

      candidates.push({
        code: normalizedCode,
        digits: normalizedCode.length,
        confidence: Number(confidence.toFixed(2)),
        patternMatched: `numeric_${digits}_${closestKeyword ? 'proximate' : 'isolated'}`,
        contextSnippet: snippet,
        rawText: rawCode,
      });
    }
  }

  // 4. Custom patterns if provided
  if (options?.customPatterns) {
    for (const customReg of options.customPatterns) {
      const matches = Array.from(rawText.matchAll(new RegExp(customReg, 'g')));
      for (const match of matches) {
        const code = match[1] || match[0];
        if (!seenCodes.has(code)) {
          seenCodes.add(code);
          candidates.push({
            code,
            digits: code.length,
            confidence: 0.9,
            patternMatched: 'custom_pattern',
            contextSnippet: match[0],
            rawText: match[0],
          });
        }
      }
    }
  }

  // Sort candidates by highest confidence first, then descending code length
  candidates.sort((a, b) => b.confidence - a.confidence || b.digits - a.digits);

  return candidates;
}
