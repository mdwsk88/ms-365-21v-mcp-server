import type { AppConfig } from './config.js';
import { GraphHttpError, graphRequest, listMessageAttachments, type QueryValue } from './graph.js';
import { getRequestContext } from './request-context.js';
import { htmlToText, normalizeText, preview, truncateText } from './text.js';

// High-frequency read scenarios: Teams chats, mail, calendar and files.
// Each function returns compact, model-oriented records instead of raw Graph payloads.

type JsonRecord = Record<string, unknown>;

function record(value: unknown): JsonRecord {
  return value && typeof value === 'object' && !Array.isArray(value) ? (value as JsonRecord) : {};
}

function records(value: unknown): JsonRecord[] {
  return Array.isArray(value) ? value.map(record) : [];
}

function text(value: unknown): string | undefined {
  return typeof value === 'string' && value ? value : undefined;
}

function segment(value: string): string {
  return encodeURIComponent(value);
}

function clamp(value: number | undefined, fallback: number, min: number, max: number): number {
  if (value === undefined || !Number.isFinite(value)) return fallback;
  return Math.max(min, Math.min(max, Math.trunc(value)));
}

function escapeOData(value: string): string {
  return value.replace(/'/g, "''");
}

/**
 * Parses an ISO 8601 date or date-time. Values without an offset are read as China Standard Time,
 * the business time zone of 21Vianet tenants. Returns a UTC ISO string for OData filters.
 */
export function toUtcInstant(value: string | undefined, name: string): string | undefined {
  if (value === undefined) return undefined;
  const trimmed = value.trim();
  if (!trimmed) return undefined;
  const dateOnly = /^\d{4}-\d{2}-\d{2}$/.test(trimmed);
  const hasOffset = /(?:Z|[+-]\d{2}:?\d{2})$/i.test(trimmed);
  const normalized = dateOnly ? `${trimmed}T00:00:00+08:00` : hasOffset ? trimmed : `${trimmed}+08:00`;
  const parsed = Date.parse(normalized);
  if (!Number.isFinite(parsed)) {
    throw new Error(`${name} must be an ISO 8601 date or date-time, for example 2026-10-12 or 2026-10-12T09:00.`);
  }
  return new Date(parsed).toISOString().replace('.000Z', 'Z');
}

export type CurrentUser = { id?: string; email?: string };

/** The signed-in user from the inbound token; empty in local stdio mode. */
export function currentUser(): CurrentUser {
  const claims = getRequestContext()?.inboundClaims;
  const email = text(claims?.preferred_username) ?? text(claims?.upn) ?? text(claims?.email);
  return { id: text(claims?.oid), email: email?.toLowerCase() };
}

/**
 * Accepts only a nextLink previously returned for the same kind of resource, so a cursor
 * cannot be used to reach a Graph path that the calling tool does not cover.
 */
export function cursorPath(config: AppConfig, cursor: string, allowed: RegExp): string {
  let url: URL;
  try {
    url = new URL(cursor);
  } catch {
    throw new Error('cursor must be the nextCursor value returned by the previous call.');
  }
  const base = new URL(config.graphBaseUrl);
  if (url.origin !== base.origin || !url.pathname.startsWith(base.pathname)) {
    throw new Error('cursor must be the nextCursor value returned by the previous call.');
  }
  const relative = url.pathname.slice(base.pathname.length) || '/';
  if (!allowed.test(decodeURIComponent(relative))) {
    throw new Error('cursor belongs to a different resource; call this tool without cursor first.');
  }
  return url.toString();
}

function nextLink(response: unknown): string | undefined {
  return text(record(response)['@odata.nextLink']);
}

function emailText(value: unknown): string | undefined {
  const address = record(record(value).emailAddress);
  const name = text(address.name);
  const mail = text(address.address);
  if (name && mail && name.toLowerCase() !== mail.toLowerCase()) return `${name} <${mail}>`;
  return mail ?? name;
}

function emailAddressOf(value: unknown): string | undefined {
  return text(record(record(value).emailAddress).address)?.toLowerCase();
}

// Teams chats

export type ChatType = 'oneOnOne' | 'group' | 'meeting';

export function shapeChat(chat: JsonRecord, me: CurrentUser) {
  const members = records(chat.members);
  const isMe = (member: JsonRecord) =>
    (me.id !== undefined && text(member.userId) === me.id) ||
    (me.email !== undefined && text(member.email)?.toLowerCase() === me.email);
  const others = members.filter(member => !isMe(member));
  const otherNames = others.map(member => text(member.displayName)).filter((name): name is string => Boolean(name));
  const last = record(chat.lastMessagePreview);
  const lastBody = record(last.body);
  const lastFrom = record(record(last.from).user);
  const lastAt = text(last.createdDateTime);
  const topic = text(chat.topic);
  const chatType = text(chat.chatType);
  return {
    id: text(chat.id),
    chatType,
    title: topic ?? (otherNames.length ? otherNames.slice(0, 6).join(', ') : chatType ?? 'chat'),
    members: otherNames.slice(0, 20),
    memberCount: members.length || undefined,
    lastMessage: lastAt
      ? {
          at: lastAt,
          from: text(lastFrom.displayName),
          fromMe: me.id !== undefined && text(lastFrom.id) === me.id ? true : undefined,
          preview: last.isDeleted === true ? '(deleted)' : preview(htmlToText(text(lastBody.content) ?? ''), 200),
          type: text(last.messageType) !== 'message' ? text(last.messageType) : undefined
        }
      : undefined,
    webUrl: text(chat.webUrl)
  };
}

function chatMatches(chat: JsonRecord, query: string): boolean {
  const needle = query.normalize('NFKC').toLowerCase().trim();
  if (!needle) return true;
  const haystack = [
    text(chat.topic),
    ...records(chat.members).flatMap(member => [text(member.displayName), text(member.email)])
  ]
    .filter(Boolean)
    .join('\n')
    .normalize('NFKC')
    .toLowerCase();
  return needle.split(/\s+/).every(part => haystack.includes(part));
}

export async function listChatsOverview(
  config: AppConfig,
  input: { top?: number; since?: string; query?: string; chatType?: ChatType; cursor?: string }
) {
  const top = clamp(input.top, 20, 1, 50);
  const since = toUtcInstant(input.since, 'since');
  const filtering = Boolean(since || input.query?.trim() || input.chatType);
  const me = currentUser();
  const maxPages = filtering ? 5 : 1;
  const results: ReturnType<typeof shapeChat>[] = [];
  let pageUrl: string | undefined = input.cursor
    ? cursorPath(config, input.cursor, /^\/(me|users\/[^/]+)\/chats$/)
    : undefined;
  let next: string | undefined;
  let reachedOlder = false;
  let truncated = false;
  const seen = new Set<string>();

  for (let page = 0; page < maxPages; page += 1) {
    const response = await graphRequest(config, 'GET', pageUrl ?? '/me/chats', {
      query: pageUrl
        ? undefined
        : {
            $top: filtering ? 50 : top,
            $select: 'id,topic,chatType,webUrl',
            $expand: 'members,lastMessagePreview',
            $orderby: 'lastMessagePreview/createdDateTime desc'
          },
      scopes: ['Chat.Read']
    });
    next = nextLink(response);
    for (const chat of records(record(response).value)) {
      const lastAt = text(record(chat.lastMessagePreview).createdDateTime);
      if (since && (!lastAt || lastAt < since)) {
        reachedOlder = true;
        break;
      }
      const id = text(chat.id);
      if (id && seen.has(id)) continue;
      if (id) seen.add(id);
      if (input.chatType && text(chat.chatType) !== input.chatType) continue;
      if (input.query && !chatMatches(chat, input.query)) continue;
      if (results.length >= top) {
        truncated = true;
        break;
      }
      results.push(shapeChat(chat, me));
    }
    if (reachedOlder || truncated || !next || results.length >= top) break;
    pageUrl = next;
  }

  return {
    count: results.length,
    chats: results,
    nextCursor: !reachedOlder && !truncated && results.length >= top ? next : undefined,
    hasMore: truncated || undefined,
    note: filtering && !reachedOlder && !truncated && next && results.length < top
      ? 'Scanned the 250 most recently active chats; narrow the query or pass since to look further.'
      : undefined
  };
}

export function shapeChatMessage(message: JsonRecord, me: CurrentUser, format: 'text' | 'html', maxChars: number) {
  const from = record(message.from);
  const user = record(from.user);
  const body = record(message.body);
  const raw = text(body.content) ?? '';
  const rendered = format === 'html' ? raw : text(body.contentType)?.toLowerCase() === 'html' ? htmlToText(raw) : normalizeText(raw);
  const { text: content, truncated } = truncateText(rendered, maxChars);
  const mentions = records(message.mentions);
  const attachments = records(message.attachments).map(attachment => {
    const contentType = text(attachment.contentType);
    if (contentType === 'messageReference') {
      const reference = record(safeJson(text(attachment.content)));
      return {
        type: 'quote',
        from: text(record(record(reference.messageSender).user).displayName),
        preview: preview(text(reference.messagePreview), 200)
      };
    }
    return {
      type: contentType === 'reference' ? 'file' : contentType,
      name: text(attachment.name),
      url: text(attachment.contentUrl)
    };
  });
  const messageType = text(message.messageType);
  return {
    id: text(message.id),
    at: text(message.createdDateTime),
    from: text(user.displayName) ?? text(record(from.application).displayName),
    fromMe: me.id !== undefined && text(user.id) === me.id ? true : undefined,
    text: messageType === 'message' ? content : content || messageType,
    truncated: truncated || undefined,
    mentions: mentions.length
      ? [...new Set(mentions.map(mention => text(mention.mentionText)).filter((value): value is string => Boolean(value)))]
      : undefined,
    mentionsMe:
      me.id !== undefined && mentions.some(mention => text(record(record(mention.mentioned).user).id) === me.id)
        ? true
        : undefined,
    attachments: attachments.length ? attachments : undefined,
    importance: text(message.importance) && text(message.importance) !== 'normal' ? text(message.importance) : undefined,
    edited: text(message.lastEditedDateTime) ? true : undefined,
    type: messageType !== 'message' ? messageType : undefined
  };
}

function safeJson(value: string | undefined): unknown {
  if (!value) return undefined;
  try {
    return JSON.parse(value);
  } catch {
    return undefined;
  }
}

export function chatMessagesQuery(top: number, since?: string, until?: string): Record<string, QueryValue> {
  // 21V Graph accepts lastModifiedDateTime gt/lt or createdDateTime lt, each with the matching
  // descending $orderby. createdDateTime gt is rejected, so the lower bound uses lastModifiedDateTime
  // and the exact window is applied locally.
  if (since) {
    return { $top: top, $orderby: 'lastModifiedDateTime desc', $filter: `lastModifiedDateTime gt ${since}` };
  }
  if (until) {
    return { $top: top, $orderby: 'createdDateTime desc', $filter: `createdDateTime lt ${until}` };
  }
  return { $top: top, $orderby: 'createdDateTime desc' };
}

export async function listChatMessagesWindow(
  config: AppConfig,
  input: {
    chatId: string;
    top?: number;
    since?: string;
    until?: string;
    cursor?: string;
    includeSystemMessages?: boolean;
    format?: 'text' | 'html';
  }
) {
  const top = clamp(input.top, 20, 1, 50);
  const since = toUtcInstant(input.since, 'since');
  const until = toUtcInstant(input.until, 'until');
  if (since && until && since >= until) throw new Error('since must be earlier than until.');
  const path = input.cursor
    ? cursorPath(config, input.cursor, /^\/chats\/[^/]+\/messages$/)
    : `/chats/${segment(input.chatId)}/messages`;
  const response = await graphRequest(config, 'GET', path, {
    query: input.cursor ? undefined : chatMessagesQuery(top, since, until),
    scopes: ['Chat.Read']
  });
  const me = currentUser();
  const format = input.format ?? 'text';
  const messages = records(record(response).value)
    .filter(message => !text(message.deletedDateTime))
    .filter(message => input.includeSystemMessages || text(message.messageType) === 'message')
    .filter(message => {
      const at = text(message.createdDateTime);
      if (!at) return false;
      return (!since || at >= since) && (!until || at < until);
    })
    .sort((left, right) => (text(left.createdDateTime) ?? '').localeCompare(text(right.createdDateTime) ?? ''))
    .map(message => shapeChatMessage(message, me, format, 4000));
  return {
    chatId: input.chatId,
    count: messages.length,
    order: 'oldest first within this page; nextCursor returns older messages',
    messages,
    nextCursor: nextLink(response)
  };
}

// Mail

const mailListSelect =
  'id,subject,from,toRecipients,ccRecipients,receivedDateTime,bodyPreview,isRead,importance,hasAttachments,flag,conversationId,webLink';

export function shapeMailRow(message: JsonRecord, me: CurrentUser) {
  const to = records(message.toRecipients);
  const cc = records(message.ccRecipients);
  const flag = record(message.flag);
  const myRole =
    me.email === undefined
      ? undefined
      : to.some(recipient => emailAddressOf(recipient) === me.email)
        ? 'to'
        : cc.some(recipient => emailAddressOf(recipient) === me.email)
          ? 'cc'
          : undefined;
  return {
    id: text(message.id),
    subject: text(message.subject) ?? '(no subject)',
    from: emailText(message.from),
    to: to.slice(0, 10).map(emailText).filter(Boolean),
    toCount: to.length > 10 ? to.length : undefined,
    ccCount: cc.length || undefined,
    myRole,
    receivedAt: text(message.receivedDateTime),
    preview: preview(text(message.bodyPreview), 300),
    isRead: message.isRead === true,
    flagged: text(flag.flagStatus) === 'flagged' ? true : undefined,
    flagDue: text(record(flag.dueDateTime).dateTime),
    importance: text(message.importance) && text(message.importance) !== 'normal' ? text(message.importance) : undefined,
    hasAttachments: message.hasAttachments === true ? true : undefined,
    conversationId: text(message.conversationId),
    webLink: text(message.webLink)
  };
}

function mailFolderPath(folder: string | undefined): string {
  const value = folder?.trim() || 'inbox';
  return value.toLowerCase() === 'all' ? '/me/messages' : `/me/mailFolders/${segment(value)}/messages`;
}

const mailCursorPattern = /^\/me\/(messages|mailFolders\/[^/]+\/messages)$/;

export function mailListFilter(input: {
  since?: string;
  until?: string;
  unreadOnly?: boolean;
  flaggedOnly?: boolean;
  from?: string;
  importance?: 'low' | 'normal' | 'high';
}): string | undefined {
  const conditions: string[] = [];
  if (input.until) conditions.push(`receivedDateTime lt ${input.until}`);
  if (input.flaggedOnly) conditions.push("flag/flagStatus eq 'flagged'");
  if (input.unreadOnly) conditions.push('isRead eq false');
  if (input.importance) conditions.push(`importance eq '${input.importance}'`);
  if (input.from) conditions.push(`from/emailAddress/address eq '${escapeOData(input.from.trim().toLowerCase())}'`);
  if (!input.since && conditions.length === 0) return undefined;
  // With $orderby=receivedDateTime, Graph requires receivedDateTime to lead the filter.
  return [`receivedDateTime ge ${input.since ?? '1900-01-01T00:00:00Z'}`, ...conditions].join(' and ');
}

export async function listMailOverview(
  config: AppConfig,
  input: {
    top?: number;
    folder?: string;
    since?: string;
    until?: string;
    unreadOnly?: boolean;
    flaggedOnly?: boolean;
    from?: string;
    importance?: 'low' | 'normal' | 'high';
    cursor?: string;
  }
) {
  const top = clamp(input.top, 20, 1, 50);
  const since = toUtcInstant(input.since, 'since');
  const until = toUtcInstant(input.until, 'until');
  if (input.from && !input.from.includes('@')) {
    throw new Error('from must be an email address. To find mail by a person name, use mail_search_messages with "from:<name>".');
  }
  const response = input.cursor
    ? await graphRequest(config, 'GET', cursorPath(config, input.cursor, mailCursorPattern), { scopes: ['Mail.Read'] })
    : await graphRequest(config, 'GET', mailFolderPath(input.folder), {
        query: {
          $top: top,
          $select: mailListSelect,
          $orderby: 'receivedDateTime desc',
          $filter: mailListFilter({ ...input, since, until })
        },
        scopes: ['Mail.Read']
      });
  const me = currentUser();
  const messages = records(record(response).value).map(message => shapeMailRow(message, me));
  return { folder: input.folder?.trim() || 'inbox', count: messages.length, messages, nextCursor: nextLink(response) };
}

export async function searchMailOverview(
  config: AppConfig,
  input: { query: string; top?: number; folder?: string; cursor?: string }
) {
  const top = clamp(input.top, 20, 1, 50);
  const response = input.cursor
    ? await graphRequest(config, 'GET', cursorPath(config, input.cursor, mailCursorPattern), { scopes: ['Mail.Read'] })
    : await graphRequest(config, 'GET', mailFolderPath(input.folder?.trim() || 'all'), {
        query: {
          $search: `"${input.query.replace(/"/g, '\\"')}"`,
          $top: top,
          $select: mailListSelect
        },
        scopes: ['Mail.Read']
      });
  const me = currentUser();
  const messages = records(record(response).value).map(message => shapeMailRow(message, me));
  return { query: input.query, count: messages.length, messages, nextCursor: nextLink(response) };
}

export async function readMailMessage(
  config: AppConfig,
  input: {
    messageId: string;
    bodyFormat?: 'text' | 'html';
    onlyNewContent?: boolean;
    maxChars?: number;
    offset?: number;
    includeAttachments?: boolean;
  }
) {
  const bodyFormat = input.bodyFormat ?? 'text';
  const bodyField = input.onlyNewContent ? 'uniqueBody' : 'body';
  const message = record(
    await graphRequest(config, 'GET', `/me/messages/${segment(input.messageId)}`, {
      query: {
        $select: `id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,sentDateTime,importance,isRead,flag,hasAttachments,conversationId,webLink,${bodyField}`
      },
      prefer: bodyFormat === 'text' ? 'outlook.body-content-type="text"' : undefined,
      scopes: ['Mail.Read']
    })
  );
  const body = record(message[bodyField]);
  const rawBody = text(body.content) ?? '';
  const fullText =
    bodyFormat === 'html'
      ? rawBody
      : text(body.contentType)?.toLowerCase() === 'html'
        ? htmlToText(rawBody)
        : normalizeText(rawBody);
  const characters = Array.from(fullText);
  const maxChars = clamp(input.maxChars, 20000, 200, 100000);
  const offset = clamp(input.offset, 0, 0, Math.max(0, characters.length));
  const slice = characters.slice(offset, offset + maxChars).join('');
  const end = offset + Array.from(slice).length;

  const attachments =
    input.includeAttachments !== false && message.hasAttachments === true
      ? records(record(await listMessageAttachments(config, input.messageId)).value).map(attachment => ({
          id: text(attachment.id),
          name: text(attachment.name),
          contentType: text(attachment.contentType),
          size: typeof attachment.size === 'number' ? attachment.size : undefined,
          inline: attachment.isInline === true ? true : undefined
        }))
      : undefined;

  const me = currentUser();
  const row = shapeMailRow(message, me);
  return {
    id: row.id,
    subject: row.subject,
    from: row.from,
    to: records(message.toRecipients).map(emailText).filter(Boolean),
    cc: records(message.ccRecipients).map(emailText).filter(Boolean),
    bcc: records(message.bccRecipients).map(emailText).filter(Boolean),
    myRole: row.myRole,
    receivedAt: row.receivedAt,
    sentAt: text(message.sentDateTime),
    isRead: row.isRead,
    flagged: row.flagged,
    flagDue: row.flagDue,
    importance: row.importance,
    conversationId: row.conversationId,
    webLink: row.webLink,
    body: {
      format: bodyFormat,
      onlyNewContent: input.onlyNewContent ? true : undefined,
      text: slice,
      offset,
      totalChars: characters.length,
      nextOffset: end < characters.length ? end : undefined
    },
    attachments,
    attachmentNote: attachments?.length
      ? 'Use mail_get_attachment through gateway_execute_tool to read attachment bytes.'
      : undefined
  };
}

// Calendar

export function shapeCalendarEvent(event: JsonRecord) {
  const attendees = records(event.attendees);
  const start = record(event.start);
  const end = record(event.end);
  const trimSeconds = (value: string | undefined) => value?.replace(/:\d{2}(\.\d+)?$/, '');
  return {
    id: text(event.id),
    subject: text(event.subject) ?? '(no subject)',
    start: trimSeconds(text(start.dateTime)),
    end: trimSeconds(text(end.dateTime)),
    isAllDay: event.isAllDay === true ? true : undefined,
    isCancelled: event.isCancelled === true ? true : undefined,
    location: text(record(event.location).displayName),
    organizer: emailText(event.organizer),
    myResponse: text(record(event.responseStatus).response),
    showAs: text(event.showAs),
    isOnlineMeeting: event.isOnlineMeeting === true ? true : undefined,
    recurring: text(event.seriesMasterId) ? true : undefined,
    attendeeCount: attendees.length || undefined,
    attendees: attendees.length
      ? attendees.slice(0, 10).map(attendee => text(record(record(attendee).emailAddress).name) ?? emailAddressOf(attendee))
      : undefined,
    preview: preview(text(event.bodyPreview), 200),
    webLink: text(event.webLink)
  };
}

export async function listCalendarOverview(
  config: AppConfig,
  input: { startDateTime: string; endDateTime: string; top?: number; timeZone?: string }
) {
  const timeZone = (input.timeZone ?? 'China Standard Time').replace(/"/g, '');
  const response = await graphRequest(config, 'GET', '/me/calendarView', {
    query: {
      startDateTime: input.startDateTime,
      endDateTime: input.endDateTime,
      $top: clamp(input.top, 50, 1, 100),
      $select:
        'id,subject,organizer,attendees,start,end,location,webLink,isOnlineMeeting,responseStatus,seriesMasterId,isAllDay,isCancelled,showAs,bodyPreview',
      $orderby: 'start/dateTime'
    },
    prefer: `outlook.timezone="${timeZone}"`,
    scopes: ['Calendars.Read']
  });
  const events = records(record(response).value).map(shapeCalendarEvent);
  return { timeZone, count: events.length, events, nextCursor: nextLink(response) };
}

// Files (OneDrive + SharePoint)

const textExtensions = new Set([
  'txt', 'md', 'markdown', 'csv', 'tsv', 'json', 'xml', 'yaml', 'yml', 'log', 'ini', 'conf', 'cfg',
  'html', 'htm', 'sql', 'py', 'js', 'ts', 'sh', 'ps1', 'java', 'cs', 'go', 'rs', 'properties'
]);

function extensionOf(name: string | undefined): string | undefined {
  const match = name?.toLowerCase().match(/\.([a-z0-9]+)$/);
  return match?.[1];
}

export function isTextLike(name: string | undefined, mimeType: string | undefined): boolean {
  const extension = extensionOf(name);
  if (extension && textExtensions.has(extension)) return true;
  const mime = mimeType?.toLowerCase() ?? '';
  return mime.startsWith('text/') || mime === 'application/json' || mime === 'application/xml';
}

export function sharingToken(url: string): string {
  const encoded = Buffer.from(url.trim(), 'utf8').toString('base64').replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');
  return `u!${encoded}`;
}

function locationOf(webUrl: string | undefined): string | undefined {
  if (!webUrl) return undefined;
  try {
    const url = new URL(webUrl);
    const parts = url.pathname.split('/').filter(Boolean).map(part => decodeURIComponent(part));
    if (url.hostname.includes('-my.')) return 'OneDrive';
    if ((parts[0] === 'sites' || parts[0] === 'teams') && parts[1]) return `SharePoint: ${parts[1]}`;
    return url.hostname;
  } catch {
    return undefined;
  }
}

export function shapeFileItem(item: JsonRecord, summary?: string) {
  const parent = record(item.parentReference);
  const lastModifiedBy = record(record(item.lastModifiedBy).user);
  const webUrl = text(item.webUrl);
  return {
    name: text(item.name),
    fileType: extensionOf(text(item.name)),
    isFolder: item.folder ? true : undefined,
    location: locationOf(webUrl),
    path: text(parent.path)?.replace(/^\/drives?\/[^/]*\/root:?/, '') || undefined,
    size: typeof item.size === 'number' ? item.size : undefined,
    lastModifiedAt: text(item.lastModifiedDateTime),
    lastModifiedBy: text(lastModifiedBy.displayName),
    summary: summary ? preview(summary.replace(/<\/?c\d+>/g, '').replace(/<ddd\/>/g, '…'), 300) : undefined,
    webUrl,
    driveId: text(parent.driveId),
    itemId: text(item.id)
  };
}

export function fileSearchKql(input: { query: string; fileTypes?: string[]; modifiedAfter?: string; siteUrl?: string }): string {
  const parts = [input.query.trim()];
  const types = (input.fileTypes ?? []).map(type => type.replace(/^\./, '').toLowerCase()).filter(type => /^[a-z0-9]+$/.test(type));
  if (types.length) parts.push(`(${types.map(type => `filetype:${type}`).join(' OR ')})`);
  if (input.modifiedAfter) parts.push(`LastModifiedTime>=${input.modifiedAfter.slice(0, 10)}`);
  if (input.siteUrl) parts.push(`path:"${input.siteUrl.replace(/"/g, '')}"`);
  return parts.join(' ');
}

function matchesFileFilters(
  item: ReturnType<typeof shapeFileItem>,
  input: { scope?: 'all' | 'onedrive' | 'sharepoint'; fileTypes?: string[]; modifiedAfter?: string }
): boolean {
  if (input.scope === 'onedrive' && item.location !== 'OneDrive') return false;
  if (input.scope === 'sharepoint' && item.location === 'OneDrive') return false;
  const types = (input.fileTypes ?? []).map(type => type.replace(/^\./, '').toLowerCase());
  if (types.length && (!item.fileType || !types.includes(item.fileType))) return false;
  if (input.modifiedAfter && (item.lastModifiedAt ?? '') < input.modifiedAfter) return false;
  return true;
}

async function resolveSite(config: AppConfig, site: string): Promise<{ id: string; webUrl?: string }> {
  const trimmed = site.trim();
  let response: unknown;
  if (/^https?:\/\//i.test(trimmed)) {
    const url = new URL(trimmed);
    const sitePath = url.pathname.split('/').filter(Boolean).slice(0, 2).map(segment).join('/');
    response = await graphRequest(config, 'GET', sitePath ? `/sites/${url.hostname}:/${sitePath}` : `/sites/${url.hostname}`, {
      query: { $select: 'id,webUrl' },
      scopes: ['Sites.Read.All']
    });
  } else {
    response = await graphRequest(config, 'GET', `/sites/${segment(trimmed)}`, {
      query: { $select: 'id,webUrl' },
      scopes: ['Sites.Read.All']
    });
  }
  const id = text(record(response).id);
  if (!id) throw new Error('Could not resolve the SharePoint site.');
  return { id, webUrl: text(record(response).webUrl) };
}

const driveItemSelect = 'id,name,webUrl,size,lastModifiedDateTime,lastModifiedBy,parentReference,file,folder';

export async function searchFiles(
  config: AppConfig,
  input: {
    query: string;
    scope?: 'all' | 'onedrive' | 'sharepoint';
    site?: string;
    fileTypes?: string[];
    modifiedAfter?: string;
    top?: number;
  }
) {
  const top = clamp(input.top, 20, 1, 50);
  const modifiedAfter = toUtcInstant(input.modifiedAfter, 'modifiedAfter');
  const site = input.site ? await resolveSite(config, input.site) : undefined;
  const filters = { scope: input.scope, fileTypes: input.fileTypes, modifiedAfter };
  const notes: string[] = [];

  try {
    const response = await graphRequest(config, 'POST', '/search/query', {
      body: {
        requests: [
          {
            entityTypes: ['driveItem'],
            query: {
              queryString: fileSearchKql({ query: input.query, fileTypes: input.fileTypes, modifiedAfter, siteUrl: site?.webUrl })
            },
            from: 0,
            size: Math.min(50, top * 2)
          }
        ]
      },
      scopes: ['Files.Read.All', 'Sites.Read.All']
    });
    const hits = records(record(records(record(records(record(response).value)[0]).hitsContainers)[0]).hits);
    const files = hits
      .map(hit => shapeFileItem(record(hit.resource), text(hit.summary)))
      .filter(item => matchesFileFilters(item, filters))
      .slice(0, top);
    return { engine: 'microsoft_search', query: input.query, count: files.length, files };
  } catch (error) {
    if (!(error instanceof GraphHttpError) || ![400, 403, 404, 405, 501].includes(error.status)) throw error;
    notes.push(`Microsoft Search is unavailable here (HTTP ${error.status}); used per-library search instead.`);
  }

  const searches: Array<Promise<unknown>> = [];
  const escaped = escapeOData(input.query.trim());
  if (site) {
    const drives = records(
      record(
        await graphRequest(config, 'GET', `/sites/${segment(site.id)}/drives`, {
          query: { $top: 10, $select: 'id,name' },
          scopes: ['Files.Read.All']
        })
      ).value
    ).slice(0, 5);
    for (const drive of drives) {
      searches.push(
        graphRequest(config, 'GET', `/drives/${segment(text(drive.id) ?? '')}/root/search(q='${escaped}')`, {
          query: { $top: top, $select: driveItemSelect },
          scopes: ['Files.Read.All']
        })
      );
    }
  } else {
    if (input.scope !== 'sharepoint') {
      searches.push(
        graphRequest(config, 'GET', `/me/drive/root/search(q='${escaped}')`, {
          query: { $top: top, $select: driveItemSelect },
          scopes: ['Files.Read']
        })
      );
    }
    if (input.scope !== 'onedrive') {
      notes.push('Without Microsoft Search, SharePoint libraries are searched only when site is provided.');
    }
  }
  const settled = await Promise.allSettled(searches);
  const failures = settled.filter(result => result.status === 'rejected').length;
  if (failures) notes.push(`${failures} library search(es) failed and were skipped.`);
  const files = settled
    .flatMap(result => (result.status === 'fulfilled' ? records(record(result.value).value) : []))
    .map(item => shapeFileItem(item))
    .filter(item => matchesFileFilters(item, filters))
    .sort((left, right) => (right.lastModifiedAt ?? '').localeCompare(left.lastModifiedAt ?? ''))
    .slice(0, top);
  return { engine: 'library_search', query: input.query, count: files.length, files, notes };
}

async function resolveDriveItem(
  config: AppConfig,
  input: { url?: string; driveId?: string; itemId?: string }
): Promise<JsonRecord> {
  if (input.url) {
    try {
      return record(
        await graphRequest(config, 'GET', `/shares/${sharingToken(input.url)}/driveItem`, {
          headers: { Prefer: 'redeemSharingLinkIfNecessary' },
          scopes: ['Files.Read.All']
        })
      );
    } catch (error) {
      if (error instanceof GraphHttpError && [400, 403, 404].includes(error.status)) {
        throw new Error(
          `Could not open this link (HTTP ${error.status}). Use a OneDrive or SharePoint file link that the signed-in user can open.`
        );
      }
      throw error;
    }
  }
  if (input.itemId && input.driveId) {
    return record(
      await graphRequest(config, 'GET', `/drives/${segment(input.driveId)}/items/${segment(input.itemId)}`, {
        scopes: ['Files.Read.All']
      })
    );
  }
  if (input.itemId) {
    return record(await graphRequest(config, 'GET', `/me/drive/items/${segment(input.itemId)}`, { scopes: ['Files.Read'] }));
  }
  throw new Error('Provide url, or itemId with optional driveId.');
}

async function redirectLocation(config: AppConfig, path: string, query?: Record<string, QueryValue>): Promise<string | undefined> {
  const response = record(
    await graphRequest(config, 'GET', path, { query, redirect: 'manual', scopes: ['Files.Read.All'] })
  );
  return text(response.location);
}

export const downloadUrlNotice =
  'Pre-authenticated, short-lived URL: download it now with a plain HTTP GET (no Authorization header). Do not share, store or log it.';

export async function getFileDownloadUrl(
  config: AppConfig,
  input: {
    url?: string;
    driveId?: string;
    itemId?: string;
    format?: 'original' | 'pdf';
    includeText?: boolean;
    maxTextChars?: number;
  }
) {
  const item = await resolveDriveItem(config, input);
  const shaped = shapeFileItem(item);
  if (shaped.isFolder) {
    throw new Error(`"${shaped.name ?? 'This item'}" is a folder. Use files_search or list the folder through gateway_search_tools.`);
  }
  const driveId = shaped.driveId;
  const itemId = shaped.itemId;
  const mimeType = text(record(item.file).mimeType);
  const format = input.format ?? 'original';
  const notes: string[] = [];
  let downloadUrl: string | undefined;

  if (format === 'pdf') {
    if (!driveId || !itemId) throw new Error('PDF conversion needs the file drive; open it by link or driveId and itemId.');
    try {
      downloadUrl = await redirectLocation(config, `/drives/${segment(driveId)}/items/${segment(itemId)}/content`, {
        format: 'pdf'
      });
    } catch (error) {
      if (error instanceof GraphHttpError) {
        throw new Error(`PDF conversion is not available for this file (HTTP ${error.status}). Use format original.`);
      }
      throw error;
    }
  } else {
    downloadUrl = text(item['@microsoft.graph.downloadUrl']);
    if (!downloadUrl && driveId && itemId) {
      downloadUrl = await redirectLocation(config, `/drives/${segment(driveId)}/items/${segment(itemId)}/content`);
    }
  }
  if (!downloadUrl) throw new Error('Microsoft Graph did not return a download URL for this file.');

  let textContent: string | undefined;
  let textTruncated: boolean | undefined;
  const size = shaped.size ?? 0;
  if (format === 'original' && input.includeText !== false && isTextLike(shaped.name, mimeType)) {
    if (size > 1024 * 1024) {
      notes.push('Text file is larger than 1 MB; download it instead of reading it inline.');
    } else {
      try {
        const response = await fetch(downloadUrl, { signal: AbortSignal.timeout(config.graphResilience.timeoutMs) });
        if (response.ok) {
          const decoded = normalizeText(await response.text());
          const result = truncateText(decoded, clamp(input.maxTextChars, 20000, 200, 100000));
          textContent = result.text;
          textTruncated = result.truncated || undefined;
        } else {
          await response.body?.cancel().catch(() => undefined);
          notes.push(`Inline text read failed with HTTP ${response.status}; download the file instead.`);
        }
      } catch {
        notes.push('Inline text read failed; download the file instead.');
      }
    }
  }

  return {
    ...shaped,
    mimeType: format === 'pdf' ? 'application/pdf' : mimeType,
    format,
    downloadUrl,
    downloadUrlNotice,
    textContent,
    textTruncated,
    notes: notes.length ? notes : undefined,
    parsingHint:
      textContent === undefined
        ? 'Download the bytes and parse them on the client. Office files can be requested again with format pdf. Files protected by sensitivity-label encryption may not be readable after download.'
        : undefined
  };
}
