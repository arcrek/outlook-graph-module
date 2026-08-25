import type { GetMessagesOptions } from '../types/mail.js';

export function escapeODataString(value: string): string {
  return value.replace(/'/g, "''");
}

export function buildODataQuery(options?: GetMessagesOptions): string {
  if (!options) {
    return '';
  }

  const params = new URLSearchParams();

  const filterClauses: string[] = [];

  if (options.filter) {
    filterClauses.push(options.filter);
  } else {
    if (options.unreadOnly) {
      filterClauses.push('isRead eq false');
    }

    if (options.subjectContains) {
      const escaped = escapeODataString(options.subjectContains);
      filterClauses.push(`contains(subject, '${escaped}')`);
    }

    if (options.fromContains) {
      const escaped = escapeODataString(options.fromContains);
      filterClauses.push(`contains(from/emailAddress/address, '${escaped}')`);
    }

    if (options.receivedAfter) {
      const isoDate =
        options.receivedAfter instanceof Date
          ? options.receivedAfter.toISOString()
          : new Date(options.receivedAfter).toISOString();
      filterClauses.push(`receivedDateTime ge ${isoDate}`);
    }
  }

  if (filterClauses.length > 0) {
    params.set('$filter', filterClauses.join(' and '));
  }

  if (options.top !== undefined && options.top !== null) {
    params.set('$top', Math.max(1, Math.min(1000, options.top)).toString());
  }

  if (options.skip !== undefined && options.skip !== null) {
    params.set('$skip', Math.max(0, options.skip).toString());
  }

  if (options.select && options.select.length > 0) {
    params.set('$select', options.select.join(','));
  }

  if (options.orderBy) {
    params.set('$orderby', options.orderBy);
  } else if (!params.has('$orderby') && !options.search) {
    params.set('$orderby', 'receivedDateTime desc');
  }

  if (options.search) {
    const escapedSearch = options.search.replace(/"/g, '\\"');
    params.set('$search', `"${escapedSearch}"`);
  }

  const queryStr = params.toString();
  return queryStr ? `?${queryStr}` : '';
}

export function getMailEndpoint(folder?: string): string {
  if (!folder) {
    return 'https://graph.microsoft.com/v1.0/me/messages';
  }
  const cleanFolder = folder.trim().toLowerCase();
  return `https://graph.microsoft.com/v1.0/me/mailFolders/${cleanFolder}/messages`;
}
