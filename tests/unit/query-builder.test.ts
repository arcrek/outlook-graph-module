import { describe, it, expect } from 'vitest';
import {
  buildODataQuery,
  escapeODataString,
  getMailEndpoint,
} from '../../ts/http/query-builder.js';

describe('query-builder', () => {
  it('should escape single quotes for OData', () => {
    expect(escapeODataString("O'Reilly")).toBe("O''Reilly");
  });

  it('should return empty string when no options provided', () => {
    expect(buildODataQuery()).toBe('');
  });

  it('should construct endpoint URLs correctly', () => {
    expect(getMailEndpoint()).toBe(
      'https://graph.microsoft.com/v1.0/me/messages'
    );
    expect(getMailEndpoint('inbox')).toBe(
      'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages'
    );
    expect(getMailEndpoint('junkemail')).toBe(
      'https://graph.microsoft.com/v1.0/me/mailFolders/junkemail/messages'
    );
  });

  it('should format filter options correctly', () => {
    const query = buildODataQuery({
      unreadOnly: true,
      subjectContains: 'Verification',
      fromContains: 'security@service.com',
      top: 5,
    });

    expect(query).toContain('%24filter=isRead+eq+false');
    expect(query).toContain('contains%28subject%2C+%27Verification%27%29');
    expect(query).toContain(
      'contains%28from%2FemailAddress%2Faddress%2C+%27security%40service.com%27%29'
    );
    expect(query).toContain('%24top=5');
    expect(query).toContain('%24orderby=receivedDateTime+desc');
  });

  it('should support date filter and custom OData filter override', () => {
    const date = new Date('2026-08-25T12:00:00.000Z');
    const query = buildODataQuery({ receivedAfter: date });
    expect(query).toContain('receivedDateTime+ge+2026-08-25T12%3A00%3A00.000Z');

    const customQuery = buildODataQuery({
      filter: "hasAttachments eq true and importance eq 'high'",
    });
    expect(customQuery).toContain(
      '%24filter=hasAttachments+eq+true+and+importance+eq+%27high%27'
    );
  });
});
