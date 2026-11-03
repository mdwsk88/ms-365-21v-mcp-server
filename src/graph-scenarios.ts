import { attachmentLinkStore, attachmentLinkUrl, attachmentLinksAvailable } from './attachment-links.js';
import type { AppConfig } from './config.js';
import { GraphHttpError, graphAccessTokenForScopes, graphRequest, type QueryValue } from './graph.js';
import { getRequestContext } from './request-context.js';
import { requireGraphScope } from './scope-policy.js';
import { cleanMailText, fileText, htmlToText, normalizeText, preview, truncateText } from './text.js';

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

/** Whose mailbox or calendar a cursor reads: the /users/{owner} segment, or undefined for /me. */
function cursorOwner(config: AppConfig, cursor: string): string | undefined {
  const basePath = new URL(config.graphBaseUrl).pathname;
  const relative = decodeURIComponent(new URL(cursor).pathname.slice(basePath.length));
  return relative.match(/^\/users\/([^/]+)\//)?.[1];
}

function isCurrentUser(value: string): boolean {
  const lower = value.toLowerCase();
  if (lower === 'me') return true;
  const me = currentUser();
  return lower === me.email || lower === me.id?.toLowerCase();
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

const channelMessagesPattern = /^\/teams\/[^/]+\/channels\/[^/]+\/messages(\/[^/]+\/replies)?$/;

/**
 * Root posts of a Teams channel, or the replies under one post, in the same compact shape as chat
 * messages. Needs ChannelMessage.Read.All, which requires admin consent.
 */
export async function listChannelMessagesOverview(
  config: AppConfig,
  input: { teamId: string; channelId: string; messageId?: string; top?: number; cursor?: string; format?: 'text' | 'html' }
) {
  const channelPath = `/teams/${segment(input.teamId)}/channels/${segment(input.channelId)}/messages`;
  const path = input.cursor
    ? cursorPath(config, input.cursor, channelMessagesPattern)
    : input.messageId
      ? `${channelPath}/${segment(input.messageId)}/replies`
      : channelPath;
  const response = await graphRequest(config, 'GET', path, {
    query: input.cursor ? undefined : { $top: clamp(input.top, 20, 1, 50) },
    scopes: ['ChannelMessage.Read.All']
  });
  const me = currentUser();
  const format = input.format ?? 'text';
  const messages = records(record(response).value)
    .filter(message => !text(message.deletedDateTime))
    .sort((left, right) => (text(left.createdDateTime) ?? '').localeCompare(text(right.createdDateTime) ?? ''))
    .map(message => ({
      ...shapeChatMessage(message, me, format, 4000),
      subject: text(message.subject),
      webLink: text(message.webUrl)
    }));
  return {
    teamId: input.teamId,
    channelId: input.channelId,
    messageId: input.messageId,
    count: messages.length,
    order: 'oldest first within this page',
    messages,
    nextCursor: nextLink(response)
  };
}

/**
 * Microsoft Search over Teams messages the user can see. Each hit names the chat, or the team and
 * channel, so the conversation can be read next with teams_list_chat_messages or
 * teams_list_channel_messages. Channel hits need ChannelMessage.Read.All as well as Chat.Read.
 */
export async function searchTeamsMessages(config: AppConfig, input: { query: string; from?: number; size?: number }) {
  const from = clamp(input.from, 0, 0, 1000);
  const size = clamp(input.size, 20, 1, 50);
  const response = await graphRequest(config, 'POST', '/search/query', {
    body: { requests: [{ entityTypes: ['chatMessage'], query: { queryString: input.query }, from, size }] },
    scopes: ['Chat.Read']
  });
  const container = record(records(record(records(record(response).value)[0]).hitsContainers)[0]);
  const messages = records(container.hits).map(hit => {
    const resource = record(hit.resource);
    const channel = record(resource.channelIdentity);
    const sender = record(resource.from);
    const summary = text(hit.summary)?.replace(/<\/?c\d+>/g, '').replace(/<ddd\/>/g, '…');
    return {
      id: text(resource.id),
      chatId: text(resource.chatId),
      teamId: text(channel.teamId),
      channelId: text(channel.channelId),
      from: text(record(sender.emailAddress).name) ?? text(record(sender.user).displayName),
      at: text(resource.createdDateTime),
      subject: text(resource.subject),
      summary: summary ? preview(htmlToText(summary), 300) : undefined,
      webLink: text(resource.webLink) ?? text(resource.webUrl)
    };
  });
  return {
    query: input.query,
    total: typeof container.total === 'number' ? container.total : undefined,
    count: messages.length,
    messages,
    nextFrom: container.moreResultsAvailable === true ? from + messages.length : undefined
  };
}

// Mail

/**
 * The mailbox a mail call reads: the signed-in user's own, or a shared mailbox the user can open,
 * which needs the Mail.Read.Shared delegated permission.
 */
export function mailboxBase(
  config: Pick<AppConfig, 'disabledGraphScopes' | 'allowedGraphScopes'>,
  mailbox: string | undefined
): { base: string; scope: 'Mail.Read' | 'Mail.Read.Shared'; mailbox?: string } {
  const value = mailbox?.trim();
  if (!value || isCurrentUser(value)) return { base: '/me', scope: 'Mail.Read' };
  requireGraphScope(config, 'Mail.Read.Shared', 'Reading a shared mailbox');
  return { base: `/users/${segment(value)}`, scope: 'Mail.Read.Shared', mailbox: value };
}

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
    preview: preview(cleanMailText(text(message.bodyPreview) ?? '').text, 300),
    isRead: message.isRead === true,
    flagged: text(flag.flagStatus) === 'flagged' ? true : undefined,
    flagDue: text(record(flag.dueDateTime).dateTime),
    importance: text(message.importance) && text(message.importance) !== 'normal' ? text(message.importance) : undefined,
    hasAttachments: message.hasAttachments === true ? true : undefined,
    conversationId: text(message.conversationId),
    webLink: text(message.webLink)
  };
}

function mailFolderPath(base: string, folder: string | undefined): string {
  const value = folder?.trim() || 'inbox';
  return value.toLowerCase() === 'all' ? `${base}/messages` : `${base}/mailFolders/${segment(value)}/messages`;
}

const mailCursorPattern = /^\/(me|users\/[^/]+)\/(messages|mailFolders\/[^/]+\/messages)$/;

async function followMailCursor(config: AppConfig, cursor: string) {
  const url = cursorPath(config, cursor, mailCursorPattern);
  const { scope, mailbox } = mailboxBase(config, cursorOwner(config, url));
  return { response: await graphRequest(config, 'GET', url, { scopes: [scope] }), mailbox };
}

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
    mailbox?: string;
    cursor?: string;
  }
) {
  const top = clamp(input.top, 20, 1, 50);
  const since = toUtcInstant(input.since, 'since');
  const until = toUtcInstant(input.until, 'until');
  if (input.from && !input.from.includes('@')) {
    throw new Error('from must be an email address. To find mail by a person name, use mail_search_messages with "from:<name>".');
  }
  let response: unknown;
  let mailbox: string | undefined;
  if (input.cursor) {
    ({ response, mailbox } = await followMailCursor(config, input.cursor));
  } else {
    const target = mailboxBase(config, input.mailbox);
    mailbox = target.mailbox;
    response = await graphRequest(config, 'GET', mailFolderPath(target.base, input.folder), {
      query: {
        $top: top,
        $select: mailListSelect,
        $orderby: 'receivedDateTime desc',
        $filter: mailListFilter({ ...input, since, until })
      },
      scopes: [target.scope]
    });
  }
  const me = currentUser();
  const messages = records(record(response).value).map(message => shapeMailRow(message, me));
  return {
    mailbox,
    folder: input.folder?.trim() || 'inbox',
    count: messages.length,
    messages,
    nextCursor: nextLink(response)
  };
}

export async function searchMailOverview(
  config: AppConfig,
  input: { query: string; top?: number; folder?: string; mailbox?: string; cursor?: string }
) {
  const top = clamp(input.top, 20, 1, 50);
  let response: unknown;
  let mailbox: string | undefined;
  if (input.cursor) {
    ({ response, mailbox } = await followMailCursor(config, input.cursor));
  } else {
    const target = mailboxBase(config, input.mailbox);
    mailbox = target.mailbox;
    response = await graphRequest(config, 'GET', mailFolderPath(target.base, input.folder?.trim() || 'all'), {
      query: {
        $search: `"${input.query.replace(/"/g, '\\"')}"`,
        $top: top,
        $select: mailListSelect
      },
      scopes: [target.scope]
    });
  }
  const me = currentUser();
  const messages = records(record(response).value).map(message => shapeMailRow(message, me));
  return { mailbox, query: input.query, count: messages.length, messages, nextCursor: nextLink(response) };
}

type ReadMailOptions = {
  bodyFormat?: 'text' | 'html';
  onlyNewContent?: boolean;
  maxChars?: number;
  offset?: number;
  includeAttachments?: boolean;
  clean?: boolean;
  mailbox?: string;
};

export async function readMailMessage(config: AppConfig, input: ReadMailOptions & { messageId: string }) {
  const bodyFormat = input.bodyFormat ?? 'text';
  const bodyField = input.onlyNewContent ? 'uniqueBody' : 'body';
  const { base, scope, mailbox } = mailboxBase(config, input.mailbox);
  const messagePath = `${base}/messages/${segment(input.messageId)}`;
  const message = record(
    await graphRequest(config, 'GET', messagePath, {
      query: {
        $select: `id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,sentDateTime,importance,isRead,flag,hasAttachments,conversationId,webLink,${bodyField}`
      },
      prefer: bodyFormat === 'text' ? 'outlook.body-content-type="text"' : undefined,
      scopes: [scope]
    })
  );
  const body = record(message[bodyField]);
  const rawBody = text(body.content) ?? '';
  const plainText =
    bodyFormat === 'html'
      ? undefined
      : text(body.contentType)?.toLowerCase() === 'html'
        ? htmlToText(rawBody)
        : normalizeText(rawBody);
  const cleaned = plainText !== undefined && input.clean !== false ? cleanMailText(plainText) : undefined;
  const fullText = cleaned?.text ?? plainText ?? rawBody;
  const characters = Array.from(fullText);
  const maxChars = clamp(input.maxChars, 20000, 200, 100000);
  const offset = clamp(input.offset, 0, 0, Math.max(0, characters.length));
  const slice = characters.slice(offset, offset + maxChars).join('');
  const end = offset + Array.from(slice).length;

  const attachments =
    input.includeAttachments !== false && message.hasAttachments === true
      ? records(
          record(
            await graphRequest(config, 'GET', `${messagePath}/attachments`, {
              query: { $select: 'id,name,contentType,size,isInline,lastModifiedDateTime' },
              scopes: [scope]
            })
          ).value
        ).map(attachment => ({
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
    mailbox,
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
      nextOffset: end < characters.length ? end : undefined,
      linksOmitted: cleaned?.linksOmitted || undefined
    },
    attachments,
    attachmentNote: attachments?.length
      ? 'Call mail_get_attachment with an attachment id to read small text attachments inline or get a single-use download URL.'
      : undefined
  };
}

const maxBatchMessages = 10;
const batchConcurrency = 4;

async function mapWithConcurrency<T, R>(items: T[], limit: number, map: (item: T) => Promise<R>): Promise<R[]> {
  const results = new Array<R>(items.length);
  let next = 0;
  const worker = async () => {
    while (next < items.length) {
      const index = next;
      next += 1;
      results[index] = await map(items[index] as T);
    }
  };
  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, worker));
  return results;
}

function itemError(error: unknown) {
  if (error instanceof GraphHttpError) {
    const graphError = record(record(error.responseBody).error);
    return { status: error.status, code: text(graphError.code), message: text(graphError.message) ?? error.message };
  }
  const code = error && typeof error === 'object' && 'code' in error ? text((error as { code?: unknown }).code) : undefined;
  return { code, message: error instanceof Error ? error.message : String(error) };
}

/**
 * Reads several messages in one call, a few at a time. Each body starts at offset 0 and is capped at
 * maxChars (8000 by default); a message that fails is returned with its error instead of failing the batch.
 */
export async function readMailMessages(config: AppConfig, input: ReadMailOptions & { messageIds: string[] }) {
  const ids = [...new Set(input.messageIds.map(id => id.trim()).filter(Boolean))];
  if (!ids.length) throw new Error('messageIds must list at least one message ID.');
  if (ids.length > maxBatchMessages) throw new Error(`messageIds accepts at most ${maxBatchMessages} message IDs.`);
  if (input.offset) throw new Error('offset works with a single messageId; continue one long message with messageId and offset.');
  const maxChars = clamp(input.maxChars, 8000, 200, 20000);
  const messages = await mapWithConcurrency(ids, batchConcurrency, async messageId => {
    try {
      return await readMailMessage(config, { ...input, messageId, offset: 0, maxChars });
    } catch (error) {
      return { id: messageId, error: itemError(error) };
    }
  });
  return {
    count: messages.length,
    failed: messages.filter(message => 'error' in message).length || undefined,
    messages
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

/**
 * The calendar a calendar call reads: the signed-in user's own, or another user's calendar shared with
 * or delegated to them, which needs the Calendars.Read.Shared delegated permission.
 */
export function calendarBase(
  config: Pick<AppConfig, 'disabledGraphScopes' | 'allowedGraphScopes'>,
  owner: string | undefined
): { base: string; scope: 'Calendars.Read' | 'Calendars.Read.Shared'; owner?: string } {
  const value = owner?.trim();
  if (!value || isCurrentUser(value)) return { base: '/me', scope: 'Calendars.Read' };
  requireGraphScope(config, 'Calendars.Read.Shared', "Reading another person's calendar");
  return { base: `/users/${segment(value)}`, scope: 'Calendars.Read.Shared', owner: value };
}

const calendarCursorPattern = /^\/(me|users\/[^/]+)\/calendarView$/;

export async function listCalendarOverview(
  config: AppConfig,
  input: {
    startDateTime: string;
    endDateTime: string;
    top?: number;
    timeZone?: string;
    calendarOwner?: string;
    cursor?: string;
  }
) {
  const timeZone = (input.timeZone ?? 'China Standard Time').replace(/"/g, '');
  const prefer = `outlook.timezone="${timeZone}"`;
  let response: unknown;
  let owner: string | undefined;
  if (input.cursor) {
    const url = cursorPath(config, input.cursor, calendarCursorPattern);
    const target = calendarBase(config, cursorOwner(config, url));
    owner = target.owner;
    response = await graphRequest(config, 'GET', url, { prefer, scopes: [target.scope] });
  } else {
    const target = calendarBase(config, input.calendarOwner);
    owner = target.owner;
    response = await graphRequest(config, 'GET', `${target.base}/calendarView`, {
      query: {
        startDateTime: input.startDateTime,
        endDateTime: input.endDateTime,
        $top: clamp(input.top, 50, 1, 100),
        $select:
          'id,subject,organizer,attendees,start,end,location,webLink,isOnlineMeeting,responseStatus,seriesMasterId,isAllDay,isCancelled,showAs,bodyPreview',
        $orderby: 'start/dateTime'
      },
      prefer,
      scopes: [target.scope]
    });
  }
  const { events, conflictCount } = markCalendarOverlaps(records(record(response).value).map(shapeCalendarEvent));
  return { calendarOwner: owner, timeZone, count: events.length, conflictCount, events, nextCursor: nextLink(response) };
}

type CalendarEventSummary = ReturnType<typeof shapeCalendarEvent>;

/**
 * Marks events that overlap in time with overlapsWith. Cancelled, all-day, declined and "free"
 * events never conflict. Times share the requested time zone, so ISO strings compare directly.
 */
export function markCalendarOverlaps(events: CalendarEventSummary[]) {
  const busy = events
    .map((event, index) => ({ event, index }))
    .filter(
      ({ event }) =>
        event.start &&
        event.end &&
        !event.isCancelled &&
        !event.isAllDay &&
        event.showAs !== 'free' &&
        event.myResponse !== 'declined'
    )
    .sort((left, right) => (left.event.start ?? '').localeCompare(right.event.start ?? ''));
  const overlaps = new Map<number, Array<{ id?: string; subject: string }>>();
  const add = (index: number, other: CalendarEventSummary) =>
    overlaps.set(index, [...(overlaps.get(index) ?? []), { id: other.id, subject: other.subject }]);
  let conflictCount = 0;
  for (let left = 0; left < busy.length; left += 1) {
    for (let right = left + 1; right < busy.length; right += 1) {
      if ((busy[right].event.start ?? '') >= (busy[left].event.end ?? '')) break;
      conflictCount += 1;
      add(busy[left].index, busy[right].event);
      add(busy[right].index, busy[left].event);
    }
  }
  return {
    conflictCount,
    events: events.map((event, index) => (overlaps.has(index) ? { ...event, overlapsWith: overlaps.get(index) } : event))
  };
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

/** Raised when the caller's App Roles do not cover the requested data; reported as permission_denied. */
export class ScenarioPermissionError extends Error {
  readonly code = 'permission_denied';

  constructor(message: string) {
    super(message);
    this.name = 'PermissionDeniedError';
  }
}

/** SharePoint libraries report driveType documentLibrary; OneDrive for Business reports business. */
function isSharePointItem(item: JsonRecord): boolean {
  const driveType = text(record(item.parentReference).driveType);
  if (driveType) return driveType === 'documentLibrary';
  return locationOf(text(item.webUrl)) !== 'OneDrive';
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
    /** False when the caller lacks the mcp.sharepoint App Role: OneDrive results only. */
    includeSharePoint?: boolean;
  }
) {
  const top = clamp(input.top, 20, 1, 50);
  const includeSharePoint = input.includeSharePoint !== false;
  if (!includeSharePoint && (input.site || input.scope === 'sharepoint')) {
    throw new ScenarioPermissionError(
      'Searching SharePoint requires the mcp.sharepoint App Role. Use scope onedrive, or ask an administrator to assign the role.'
    );
  }
  const scope = includeSharePoint ? input.scope : 'onedrive';
  const modifiedAfter = toUtcInstant(input.modifiedAfter, 'modifiedAfter');
  const site = input.site ? await resolveSite(config, input.site) : undefined;
  const filters = { scope, fileTypes: input.fileTypes, modifiedAfter };
  const notes: string[] = [];
  if (!includeSharePoint && input.scope !== 'onedrive') {
    notes.push('SharePoint files are excluded because the signed-in user does not have the mcp.sharepoint App Role.');
  }

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
            // A one-sided scope is filtered locally, so fetch the most hits to keep enough after filtering.
            size: scope === 'onedrive' || scope === 'sharepoint' ? 50 : Math.min(50, top * 2)
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
    return {
      engine: 'microsoft_search',
      query: input.query,
      count: files.length,
      files,
      notes: notes.length ? notes : undefined
    };
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
    if (scope !== 'sharepoint') {
      searches.push(
        graphRequest(config, 'GET', `/me/drive/root/search(q='${escaped}')`, {
          query: { $top: top, $select: driveItemSelect },
          scopes: ['Files.Read']
        })
      );
    }
    if (scope !== 'onedrive') {
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
    // The shares API documents Files.ReadWrite as its least-privileged delegated permission.
    requireGraphScope(config, 'Files.ReadWrite', 'Opening a OneDrive or SharePoint file by link');
    try {
      return record(
        await graphRequest(config, 'GET', `/shares/${sharingToken(input.url)}/driveItem`, {
          headers: { Prefer: 'redeemSharingLinkIfNecessary' },
          scopes: ['Files.ReadWrite']
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

/** Microsoft Graph converts files to other formats only up to this size. */
const pdfConversionMaxBytes = 2 * 1024 * 1024;

const maxInlineTextBytes = 1024 * 1024;
const maxBase64AttachmentBytes = 10 * 1024 * 1024;

export const attachmentLinkNotice =
  'Single-use, short-lived URL served by this MCP server: download it now with a plain HTTP GET (no Authorization header). It stops working after one download. Do not share, store or log it.';

/**
 * One mail attachment for an agent: small text attachments inline, cloud-file links as their source URL,
 * everything else as a single-use download URL (or base64 when the deployment has no HTTP route).
 */
export async function getMailAttachment(
  config: AppConfig,
  input: {
    messageId: string;
    attachmentId: string;
    mailbox?: string;
    format?: 'url' | 'base64';
    includeText?: boolean;
    maxTextChars?: number;
  }
) {
  const { base, scope } = mailboxBase(config, input.mailbox);
  const attachmentPath = `${base}/messages/${segment(input.messageId)}/attachments/${segment(input.attachmentId)}`;
  const meta = record(
    await graphRequest(config, 'GET', attachmentPath, {
      query: { $select: 'id,name,contentType,size,isInline,lastModifiedDateTime' },
      scopes: [scope]
    })
  );
  const odataType = text(meta['@odata.type']) ?? '';
  const kind = odataType.endsWith('referenceAttachment') ? 'link' : odataType.endsWith('itemAttachment') ? 'item' : 'file';
  const size = typeof meta.size === 'number' ? meta.size : undefined;
  let name = text(meta.name) ?? 'attachment';
  let contentType = text(meta.contentType) ?? 'application/octet-stream';
  if (kind === 'item') {
    contentType = 'message/rfc822';
    if (!/\.[a-z0-9]{2,5}$/i.test(name)) name = `${name}.eml`;
  }
  const summary = {
    id: text(meta.id) ?? input.attachmentId,
    name,
    kind,
    contentType,
    size,
    inline: meta.isInline === true ? true : undefined
  };

  if (kind === 'link') {
    const full = record(await graphRequest(config, 'GET', attachmentPath, { scopes: [scope] }));
    return {
      ...summary,
      sourceUrl: text(full.sourceUrl),
      hint: 'This attachment links to a cloud file. Call files_get_download_url with url set to sourceUrl.'
    };
  }

  const valuePath = `${attachmentPath}/$value`;
  if (kind === 'file' && input.includeText !== false && isTextLike(name, contentType) && (size ?? 0) <= maxInlineTextBytes) {
    const bytes = record(
      await graphRequest(config, 'GET', valuePath, {
        scopes: [scope],
        responseType: 'base64',
        maxResponseBytes: maxInlineTextBytes
      })
    );
    const decoded = fileText(Buffer.from(text(bytes.contentBase64) ?? '', 'base64').toString('utf8'));
    const result = truncateText(decoded, clamp(input.maxTextChars, 20000, 200, 100000));
    return { ...summary, textContent: result.text, textTruncated: result.truncated || undefined };
  }

  const format = input.format ?? (attachmentLinksAvailable(config) ? 'url' : 'base64');
  if (format === 'base64') {
    if ((size ?? 0) > maxBase64AttachmentBytes) {
      throw new Error('Attachments larger than 10 MB are only available as a download URL (format url).');
    }
    const bytes = record(
      await graphRequest(config, 'GET', valuePath, {
        scopes: [scope],
        responseType: 'base64',
        maxResponseBytes: maxBase64AttachmentBytes
      })
    );
    return { ...summary, contentBase64: text(bytes.contentBase64) };
  }
  if (!attachmentLinksAvailable(config)) {
    throw new Error('Download URLs need the HTTP transport. Use format base64 for attachments up to 10 MB.');
  }

  const accessToken = await graphAccessTokenForScopes(config, [scope]);
  const { ticket, expiresAt } = attachmentLinkStore.create(
    { graphPath: valuePath, accessToken, name, contentType, userKey: currentUser().id ?? 'local' },
    config.attachmentUrlTtlSeconds
  );
  return {
    ...summary,
    downloadUrl: attachmentLinkUrl(config, ticket),
    expiresAt: new Date(expiresAt).toISOString(),
    downloadUrlNotice: attachmentLinkNotice,
    parsingHint:
      'Download the bytes and parse them on the client. Attachments protected by sensitivity-label encryption may not be readable after download.'
  };
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
    /** False when the caller lacks the mcp.sharepoint App Role: SharePoint files are refused. */
    includeSharePoint?: boolean;
  }
) {
  const item = await resolveDriveItem(config, input);
  if (input.includeSharePoint === false && isSharePointItem(item)) {
    throw new ScenarioPermissionError(
      'This file is stored in SharePoint. Reading SharePoint files requires the mcp.sharepoint App Role; ask an administrator to assign it.'
    );
  }
  const shaped = shapeFileItem(item);
  if (shaped.isFolder) {
    throw new Error(`"${shaped.name ?? 'This item'}" is a folder. Use files_search or list the folder through gateway_search_tools.`);
  }
  const driveId = shaped.driveId;
  const itemId = shaped.itemId;
  const mimeType = text(record(item.file).mimeType);
  const notes: string[] = [];
  let format = input.format ?? 'original';
  if (format === 'pdf' && (shaped.size ?? 0) > pdfConversionMaxBytes) {
    notes.push('PDF conversion supports files up to 2 MB, so the original file is returned instead.');
    format = 'original';
  }
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
          const decoded = fileText(await response.text());
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
