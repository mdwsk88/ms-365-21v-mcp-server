import type { AppConfig } from './config.js';
import { getGraphResilience } from './graph-resilience.js';
import { getGraphAccessToken } from './oauth.js';

export class GraphHttpError extends Error {
  readonly status: number;
  readonly responseBody: unknown;

  constructor(status: number, responseBody: unknown) {
    super(`Microsoft Graph request failed with HTTP ${status}`);
    this.name = 'GraphHttpError';
    this.status = status;
    this.responseBody = responseBody;
  }
}

export type QueryValue = string | number | boolean | undefined;
export type GraphMethod = 'GET' | 'POST' | 'PATCH' | 'PUT' | 'DELETE';

export type GraphRequestOptions = {
  query?: Record<string, QueryValue>;
  body?: unknown;
  contentType?: string;
  prefer?: string;
  scopes?: string[];
  headers?: Record<string, string>;
  responseType?: 'auto' | 'base64';
  maxResponseBytes?: number;
};

type RecipientInput = {
  email: string;
  name?: string;
};

type EventDateTimeInput = {
  dateTime: string;
  timeZone?: string;
};

type AttendeeInput = {
  email: string;
  name?: string;
  type?: 'required' | 'optional' | 'resource';
};

type ContactEmailInput = {
  address: string;
  name?: string;
};

function appendQuery(url: URL, query?: Record<string, QueryValue>): void {
  if (!query) return;
  for (const [key, value] of Object.entries(query)) {
    if (value !== undefined) url.searchParams.set(key, String(value));
  }
}

function resolveGraphUrl(config: AppConfig, pathOrUrl: string, query?: Record<string, QueryValue>): URL {
  const trimmed = pathOrUrl.trim();
  const graphBaseUrl = new URL(config.graphBaseUrl);
  let url: URL;

  if (/^https:\/\//i.test(trimmed)) {
    url = new URL(trimmed);
    if (url.origin !== graphBaseUrl.origin || !url.pathname.startsWith(graphBaseUrl.pathname)) {
      throw new Error(`Only URLs under ${config.graphBaseUrl} are allowed.`);
    }
  } else {
    const path = trimmed.startsWith('/') ? trimmed : `/${trimmed}`;
    url = new URL(`${config.graphBaseUrl}${path}`);
  }

  appendQuery(url, query);
  return url;
}

function graphScopes(config: AppConfig, scopeNames: string[]): string[] {
  return scopeNames.map(scope =>
    scope.startsWith('https://') || scope.startsWith('api://') ? scope : `${config.graphResource}/${scope}`
  );
}

export async function graphRequest(
  config: AppConfig,
  method: GraphMethod,
  pathOrUrl: string,
  options: GraphRequestOptions = {}
): Promise<unknown> {
  const accessToken = await getGraphAccessToken(config, options.scopes ? graphScopes(config, options.scopes) : undefined);
  const url = resolveGraphUrl(config, pathOrUrl, options.query);
  const headers: Record<string, string> = {
    Authorization: `Bearer ${accessToken}`,
    Accept: options.responseType === 'base64' ? '*/*' : 'application/json',
    ...options.headers
  };
  let body: BodyInit | undefined;

  if (options.prefer) headers.Prefer = options.prefer;
  if (options.body !== undefined) {
    if (typeof options.body === 'string') {
      body = options.body;
      headers['Content-Type'] = options.contentType ?? 'application/octet-stream';
    } else if (options.body instanceof Uint8Array) {
      const bytes = new Uint8Array(options.body.byteLength);
      bytes.set(options.body);
      body = new Blob([bytes.buffer]);
      headers['Content-Type'] = options.contentType ?? 'application/octet-stream';
    } else {
      body = JSON.stringify(options.body);
      headers['Content-Type'] = options.contentType ?? 'application/json';
    }
  }

  const response = await getGraphResilience(config).execute(
    signal => fetch(url, { method, headers, body, signal }),
    { retryTransientFailures: method !== 'POST' }
  );
  if (response.status === 204 || response.status === 202) {
    return {
      ok: true,
      status: response.status,
      location: response.headers.get('location') ?? undefined
    };
  }

  if (options.responseType === 'base64') {
    if (!response.ok) {
      const contentType = response.headers.get('content-type') ?? '';
      const bodyText = await response.text();
      const responseBody = contentType.includes('application/json') && bodyText ? JSON.parse(bodyText) : bodyText;
      throw new GraphHttpError(response.status, responseBody);
    }

    const maxBytes = options.maxResponseBytes ?? 10 * 1024 * 1024;
    const contentLength = Number.parseInt(response.headers.get('content-length') ?? '', 10);
    if (Number.isFinite(contentLength) && contentLength > maxBytes) {
      throw new Error(`Microsoft Graph response is ${contentLength} bytes; the MCP download limit is ${maxBytes} bytes.`);
    }
    const bytes = new Uint8Array(await response.arrayBuffer());
    if (bytes.byteLength > maxBytes) {
      throw new Error(`Microsoft Graph response is ${bytes.byteLength} bytes; the MCP download limit is ${maxBytes} bytes.`);
    }
    return {
      contentType: response.headers.get('content-type') ?? 'application/octet-stream',
      size: bytes.byteLength,
      contentBase64: Buffer.from(bytes).toString('base64')
    };
  }

  const contentType = response.headers.get('content-type') ?? '';
  const bodyText = await response.text();
  const responseBody = contentType.includes('application/json') && bodyText ? JSON.parse(bodyText) : bodyText;

  if (!response.ok) {
    throw new GraphHttpError(response.status, responseBody);
  }

  return responseBody;
}

function clampTop(top: number | undefined, fallback: number, max: number): number {
  if (!Number.isFinite(top)) return fallback;
  return Math.max(1, Math.min(max, Math.trunc(top as number)));
}

function escapeODataString(value: string): string {
  return value.replace(/'/g, "''");
}

function pathSegment(value: string): string {
  return encodeURIComponent(value);
}

function drivePath(value: string): string {
  return value
    .replace(/^\/+|\/+$/g, '')
    .split('/')
    .filter(Boolean)
    .map(segment => encodeURIComponent(segment))
    .join('/');
}

function sharePointFieldsExpand(fieldNames?: string[]): string {
  const names = fieldNames?.map(value => value.trim()).filter(Boolean) ?? [];
  return names.length ? `fields(select=${names.join(',')})` : 'fields';
}

function htmlBody(content: string) {
  return { contentType: 'HTML', content };
}

function textBody(content: string) {
  return { contentType: 'Text', content };
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function teamsBody(content: string, contentIsHtml?: boolean) {
  return {
    contentType: 'html',
    content: contentIsHtml ? content : escapeHtml(content).replace(/\r?\n/g, '<br>')
  };
}

function recipient(recipientInput: RecipientInput) {
  return {
    emailAddress: {
      address: recipientInput.email,
      name: recipientInput.name
    }
  };
}

function recipients(values: RecipientInput[] | undefined) {
  return values?.map(recipient) ?? [];
}

function attendee(value: AttendeeInput) {
  return {
    emailAddress: {
      address: value.email,
      name: value.name
    },
    type: value.type ?? 'required'
  };
}

function eventDateTime(value: EventDateTimeInput) {
  return {
    dateTime: value.dateTime,
    timeZone: value.timeZone ?? 'China Standard Time'
  };
}

function base64ToBytes(value: string): Uint8Array {
  return Uint8Array.from(Buffer.from(value, 'base64'));
}

function splitSearchWords(value: string): string[] {
  return value
    .trim()
    .split(/\s+/)
    .map(item => item.trim())
    .filter(Boolean)
    .slice(0, 5);
}

export async function graphGet(config: AppConfig, pathOrUrl: string, query?: Record<string, QueryValue>): Promise<unknown> {
  return graphRequest(config, 'GET', pathOrUrl, { query });
}

export async function getMe(config: AppConfig): Promise<unknown> {
  return graphRequest(config, 'GET', '/me', {
    query: {
      $select: 'id,displayName,userPrincipalName,mail,jobTitle,department,officeLocation,mobilePhone,businessPhones'
    },
    scopes: ['User.Read']
  });
}

export async function searchUsers(config: AppConfig, query: string, top?: number): Promise<unknown> {
  const words = splitSearchWords(query);
  const filter = words
    .map(word => {
      const escaped = escapeODataString(word);
      return `(startswith(displayName,'${escaped}') or startswith(mail,'${escaped}') or startswith(userPrincipalName,'${escaped}'))`;
    })
    .join(' or ');

  return graphRequest(config, 'GET', '/users', {
    query: {
      $top: clampTop(top, 10, 25),
      $select: 'id,displayName,userPrincipalName,mail,jobTitle,department,officeLocation',
      $filter: filter || undefined,
      $orderby: filter ? undefined : 'displayName'
    },
    scopes: ['User.ReadBasic.All']
  });
}

export async function getUser(config: AppConfig, userIdOrPrincipalName: string): Promise<unknown> {
  return graphRequest(config, 'GET', `/users/${pathSegment(userIdOrPrincipalName)}`, {
    query: {
      $select: 'id,displayName,userPrincipalName,mail,jobTitle,department,officeLocation,mobilePhone,businessPhones'
    },
    scopes: ['User.Read.All']
  });
}

export async function listMyMessages(config: AppConfig, top?: number, unreadOnly?: boolean): Promise<unknown> {
  return graphRequest(config, 'GET', '/me/messages', {
    query: {
      $top: clampTop(top, 10, 50),
      $select: 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,webLink,isRead,importance,hasAttachments',
      $orderby: 'receivedDateTime desc',
      $filter: unreadOnly ? 'isRead eq false' : undefined
    },
    scopes: ['Mail.Read']
  });
}

export async function listMyMessagesSince(config: AppConfig, since: Date, top = 50): Promise<unknown> {
  return graphRequest(config, 'GET', '/me/messages', {
    query: {
      $top: clampTop(top, 50, 50),
      $select:
        'id,subject,from,receivedDateTime,webLink,isRead,importance,hasAttachments,flag',
      $orderby: 'receivedDateTime desc',
      $filter: `receivedDateTime ge ${since.toISOString()}`
    },
    scopes: ['Mail.Read']
  });
}

export async function searchMyMessages(config: AppConfig, query: string, top?: number): Promise<unknown> {
  return graphRequest(config, 'GET', '/me/messages', {
    query: {
      $search: `"${query.replace(/"/g, '\\"')}"`,
      $top: clampTop(top, 10, 50),
      $select: 'id,subject,from,receivedDateTime,webLink,isRead,importance,hasAttachments'
    },
    scopes: ['Mail.Read']
  });
}

export async function getMessage(config: AppConfig, messageId: string): Promise<unknown> {
  return graphRequest(config, 'GET', `/me/messages/${pathSegment(messageId)}`, {
    query: {
      $select:
        'id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,sentDateTime,bodyPreview,body,webLink,isRead,importance,hasAttachments'
    },
    scopes: ['Mail.Read']
  });
}

export async function listMailFolders(config: AppConfig, top?: number): Promise<unknown> {
  return graphRequest(config, 'GET', '/me/mailFolders', {
    query: {
      $top: clampTop(top, 20, 100),
      $select: 'id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount'
    },
    scopes: ['Mail.Read']
  });
}

export async function listMessageAttachments(config: AppConfig, messageId: string): Promise<unknown> {
  return graphRequest(config, 'GET', `/me/messages/${pathSegment(messageId)}/attachments`, {
    query: {
      $select: 'id,name,contentType,size,isInline,lastModifiedDateTime'
    },
    scopes: ['Mail.Read']
  });
}

export async function sendMail(
  config: AppConfig,
  input: {
    subject: string;
    body: string;
    to: RecipientInput[];
    cc?: RecipientInput[];
    bcc?: RecipientInput[];
    saveToSentItems?: boolean;
    bodyIsHtml?: boolean;
  }
): Promise<unknown> {
  return graphRequest(config, 'POST', '/me/sendMail', {
    body: {
      message: {
        subject: input.subject,
        body: input.bodyIsHtml ? htmlBody(input.body) : textBody(input.body),
        toRecipients: recipients(input.to),
        ccRecipients: recipients(input.cc),
        bccRecipients: recipients(input.bcc)
      },
      saveToSentItems: input.saveToSentItems ?? true
    },
    scopes: ['Mail.Send']
  });
}

export async function createDraftMessage(
  config: AppConfig,
  input: {
    subject: string;
    body: string;
    to?: RecipientInput[];
    cc?: RecipientInput[];
    bcc?: RecipientInput[];
    bodyIsHtml?: boolean;
  }
): Promise<unknown> {
  return graphRequest(config, 'POST', '/me/messages', {
    body: {
      subject: input.subject,
      body: input.bodyIsHtml ? htmlBody(input.body) : textBody(input.body),
      toRecipients: recipients(input.to),
      ccRecipients: recipients(input.cc),
      bccRecipients: recipients(input.bcc)
    },
    scopes: ['Mail.ReadWrite']
  });
}

export async function sendDraftMessage(config: AppConfig, messageId: string): Promise<unknown> {
  return graphRequest(config, 'POST', `/me/messages/${pathSegment(messageId)}/send`, {
    scopes: ['Mail.Send']
  });
}

export async function replyToMessage(config: AppConfig, messageId: string, comment: string): Promise<unknown> {
  return graphRequest(config, 'POST', `/me/messages/${pathSegment(messageId)}/reply`, {
    body: { comment },
    scopes: ['Mail.Send']
  });
}

export async function forwardMessage(
  config: AppConfig,
  messageId: string,
  comment: string,
  to: RecipientInput[]
): Promise<unknown> {
  return graphRequest(config, 'POST', `/me/messages/${pathSegment(messageId)}/forward`, {
    body: {
      comment,
      toRecipients: recipients(to)
    },
    scopes: ['Mail.Send']
  });
}

export async function updateMessageReadState(config: AppConfig, messageId: string, isRead: boolean): Promise<unknown> {
  return graphRequest(config, 'PATCH', `/me/messages/${pathSegment(messageId)}`, {
    body: { isRead },
    scopes: ['Mail.ReadWrite']
  });
}

export async function moveMessage(config: AppConfig, messageId: string, destinationFolderId: string): Promise<unknown> {
  return graphRequest(config, 'POST', `/me/messages/${pathSegment(messageId)}/move`, {
    body: { destinationId: destinationFolderId },
    scopes: ['Mail.ReadWrite']
  });
}

export async function deleteMessage(config: AppConfig, messageId: string): Promise<unknown> {
  return graphRequest(config, 'DELETE', `/me/messages/${pathSegment(messageId)}`, {
    scopes: ['Mail.ReadWrite']
  });
}

export async function listMyCalendarEvents(config: AppConfig, daysAhead?: number, top?: number): Promise<unknown> {
  const days = clampTop(daysAhead, 7, 30);
  const start = new Date();
  const end = new Date(start.getTime() + days * 24 * 60 * 60 * 1000);

  return graphRequest(config, 'GET', '/me/calendarView', {
    query: {
      startDateTime: start.toISOString(),
      endDateTime: end.toISOString(),
      $top: clampTop(top, 10, 50),
      $select: 'id,subject,organizer,start,end,location,webLink,isOnlineMeeting,responseStatus',
      $orderby: 'start/dateTime'
    },
    prefer: 'outlook.timezone="China Standard Time"',
    scopes: ['Calendars.Read']
  });
}

export async function listMyCalendars(config: AppConfig, top?: number): Promise<unknown> {
  return graphRequest(config, 'GET', '/me/calendars', {
    query: {
      $top: clampTop(top, 20, 100),
      $select: 'id,name,canShare,canEdit,canViewPrivateItems,owner'
    },
    scopes: ['Calendars.Read']
  });
}

export async function getCalendarEvent(config: AppConfig, eventId: string): Promise<unknown> {
  return graphRequest(config, 'GET', `/me/events/${pathSegment(eventId)}`, {
    query: {
      $select: 'id,subject,bodyPreview,body,organizer,attendees,start,end,location,webLink,isOnlineMeeting,responseStatus'
    },
    scopes: ['Calendars.Read']
  });
}

export async function createCalendarEvent(
  config: AppConfig,
  input: {
    subject: string;
    start: EventDateTimeInput;
    end: EventDateTimeInput;
    body?: string;
    location?: string;
    attendees?: AttendeeInput[];
    isOnlineMeeting?: boolean;
  }
): Promise<unknown> {
  return graphRequest(config, 'POST', '/me/events', {
    body: {
      subject: input.subject,
      body: input.body ? htmlBody(input.body) : undefined,
      start: eventDateTime(input.start),
      end: eventDateTime(input.end),
      location: input.location ? { displayName: input.location } : undefined,
      attendees: input.attendees?.map(attendee),
      isOnlineMeeting: input.isOnlineMeeting,
      onlineMeetingProvider: input.isOnlineMeeting ? 'teamsForBusiness' : undefined
    },
    prefer: 'outlook.timezone="China Standard Time"',
    scopes: ['Calendars.ReadWrite']
  });
}

export async function updateCalendarEvent(
  config: AppConfig,
  eventId: string,
  input: {
    subject?: string;
    start?: EventDateTimeInput;
    end?: EventDateTimeInput;
    body?: string;
    location?: string;
    attendees?: AttendeeInput[];
    isOnlineMeeting?: boolean;
  }
): Promise<unknown> {
  return graphRequest(config, 'PATCH', `/me/events/${pathSegment(eventId)}`, {
    body: {
      subject: input.subject,
      body: input.body ? htmlBody(input.body) : undefined,
      start: input.start ? eventDateTime(input.start) : undefined,
      end: input.end ? eventDateTime(input.end) : undefined,
      location: input.location ? { displayName: input.location } : undefined,
      attendees: input.attendees?.map(attendee),
      isOnlineMeeting: input.isOnlineMeeting,
      onlineMeetingProvider: input.isOnlineMeeting ? 'teamsForBusiness' : undefined
    },
    prefer: 'outlook.timezone="China Standard Time"',
    scopes: ['Calendars.ReadWrite']
  });
}

export async function deleteCalendarEvent(config: AppConfig, eventId: string): Promise<unknown> {
  return graphRequest(config, 'DELETE', `/me/events/${pathSegment(eventId)}`, {
    scopes: ['Calendars.ReadWrite']
  });
}

export async function respondToCalendarEvent(
  config: AppConfig,
  eventId: string,
  response: 'accept' | 'tentativelyAccept' | 'decline',
  comment?: string,
  sendResponse = true
): Promise<unknown> {
  return graphRequest(config, 'POST', `/me/events/${pathSegment(eventId)}/${response}`, {
    body: { comment, sendResponse },
    scopes: ['Calendars.ReadWrite']
  });
}

export async function listMyDriveRoot(config: AppConfig, top?: number): Promise<unknown> {
  return graphRequest(config, 'GET', '/me/drive/root/children', {
    query: {
      $top: clampTop(top, 20, 100),
      $select: 'id,name,webUrl,size,lastModifiedDateTime,folder,file,parentReference'
    },
    scopes: ['Files.Read']
  });
}

export async function searchMyDrive(config: AppConfig, query: string, top?: number): Promise<unknown> {
  const escaped = escapeODataString(query.trim());
  return graphRequest(config, 'GET', `/me/drive/root/search(q='${escaped}')`, {
    query: {
      $top: clampTop(top, 10, 50),
      $select: 'id,name,webUrl,size,lastModifiedDateTime,folder,file,parentReference'
    },
    scopes: ['Files.Read']
  });
}

export async function getDriveItem(config: AppConfig, itemId: string): Promise<unknown> {
  return graphRequest(config, 'GET', `/me/drive/items/${pathSegment(itemId)}`, {
    query: {
      $select: 'id,name,webUrl,size,lastModifiedDateTime,folder,file,parentReference'
    },
    scopes: ['Files.Read']
  });
}

export async function createDriveFolder(
  config: AppConfig,
  name: string,
  parentItemId?: string,
  conflictBehavior: 'rename' | 'replace' | 'fail' = 'rename'
): Promise<unknown> {
  const parentPath = parentItemId ? `/me/drive/items/${pathSegment(parentItemId)}/children` : '/me/drive/root/children';
  return graphRequest(config, 'POST', parentPath, {
    body: {
      name,
      folder: {},
      '@microsoft.graph.conflictBehavior': conflictBehavior
    },
    scopes: ['Files.ReadWrite']
  });
}

export async function uploadSmallDriveFile(
  config: AppConfig,
  path: string,
  content: string,
  contentEncoding: 'text' | 'base64' = 'text',
  contentType?: string,
  conflictBehavior: 'replace' | 'rename' | 'fail' = 'replace'
): Promise<unknown> {
  const bytes = contentEncoding === 'base64' ? base64ToBytes(content) : content;
  return graphRequest(config, 'PUT', `/me/drive/root:/${drivePath(path)}:/content`, {
    query: { '@microsoft.graph.conflictBehavior': conflictBehavior },
    body: bytes,
    contentType: contentType ?? (contentEncoding === 'text' ? 'text/plain; charset=utf-8' : 'application/octet-stream'),
    scopes: ['Files.ReadWrite']
  });
}

export async function renameDriveItem(config: AppConfig, itemId: string, newName: string): Promise<unknown> {
  return graphRequest(config, 'PATCH', `/me/drive/items/${pathSegment(itemId)}`, {
    body: { name: newName },
    scopes: ['Files.ReadWrite']
  });
}

export async function moveDriveItem(config: AppConfig, itemId: string, newParentItemId: string): Promise<unknown> {
  return graphRequest(config, 'PATCH', `/me/drive/items/${pathSegment(itemId)}`, {
    body: { parentReference: { id: newParentItemId } },
    scopes: ['Files.ReadWrite']
  });
}

export async function deleteDriveItem(config: AppConfig, itemId: string): Promise<unknown> {
  return graphRequest(config, 'DELETE', `/me/drive/items/${pathSegment(itemId)}`, {
    scopes: ['Files.ReadWrite']
  });
}

export async function createDriveItemLink(
  config: AppConfig,
  itemId: string,
  type: 'view' | 'edit' | 'embed' = 'view',
  scope: 'anonymous' | 'organization' | 'users' = 'organization'
): Promise<unknown> {
  return graphRequest(config, 'POST', `/me/drive/items/${pathSegment(itemId)}/createLink`, {
    body: { type, scope },
    scopes: ['Files.ReadWrite']
  });
}

export async function listDriveItemPermissions(config: AppConfig, itemId: string): Promise<unknown> {
  return graphRequest(config, 'GET', `/me/drive/items/${pathSegment(itemId)}/permissions`, {
    scopes: ['Files.Read']
  });
}

export async function searchSharePointSites(config: AppConfig, query: string, top?: number): Promise<unknown> {
  return graphRequest(config, 'GET', '/sites', {
    query: {
      search: query.trim(),
      $top: clampTop(top, 10, 50),
      $select: 'id,name,displayName,description,webUrl,createdDateTime,lastModifiedDateTime,siteCollection'
    },
    scopes: ['Sites.Read.All']
  });
}

export async function getSharePointSite(config: AppConfig, siteId: string): Promise<unknown> {
  return graphRequest(config, 'GET', `/sites/${pathSegment(siteId)}`, {
    query: {
      $select: 'id,name,displayName,description,webUrl,createdDateTime,lastModifiedDateTime,siteCollection'
    },
    scopes: ['Sites.Read.All']
  });
}

export async function listSharePointLists(config: AppConfig, siteId: string, top?: number): Promise<unknown> {
  return graphRequest(config, 'GET', `/sites/${pathSegment(siteId)}/lists`, {
    query: {
      $top: clampTop(top, 25, 100),
      $select: 'id,name,displayName,description,webUrl,createdDateTime,lastModifiedDateTime,list'
    },
    scopes: ['Sites.Read.All']
  });
}

export async function listSharePointListItems(
  config: AppConfig,
  siteId: string,
  listId: string,
  top?: number,
  fieldNames?: string[],
  filter?: string
): Promise<unknown> {
  return graphRequest(config, 'GET', `/sites/${pathSegment(siteId)}/lists/${pathSegment(listId)}/items`, {
    query: {
      $top: clampTop(top, 25, 100),
      $expand: sharePointFieldsExpand(fieldNames),
      $filter: filter?.trim() || undefined
    },
    scopes: ['Sites.Read.All']
  });
}

export async function getSharePointListItem(
  config: AppConfig,
  siteId: string,
  listId: string,
  itemId: string,
  fieldNames?: string[]
): Promise<unknown> {
  return graphRequest(
    config,
    'GET',
    `/sites/${pathSegment(siteId)}/lists/${pathSegment(listId)}/items/${pathSegment(itemId)}`,
    {
      query: { $expand: sharePointFieldsExpand(fieldNames) },
      scopes: ['Sites.Read.All']
    }
  );
}

export async function createSharePointListItem(
  config: AppConfig,
  siteId: string,
  listId: string,
  fields: Record<string, unknown>
): Promise<unknown> {
  return graphRequest(config, 'POST', `/sites/${pathSegment(siteId)}/lists/${pathSegment(listId)}/items`, {
    body: { fields },
    scopes: ['Sites.ReadWrite.All']
  });
}

export async function updateSharePointListItem(
  config: AppConfig,
  siteId: string,
  listId: string,
  itemId: string,
  fields: Record<string, unknown>
): Promise<unknown> {
  return graphRequest(
    config,
    'PATCH',
    `/sites/${pathSegment(siteId)}/lists/${pathSegment(listId)}/items/${pathSegment(itemId)}/fields`,
    {
      body: fields,
      scopes: ['Sites.ReadWrite.All']
    }
  );
}

export async function deleteSharePointListItem(
  config: AppConfig,
  siteId: string,
  listId: string,
  itemId: string
): Promise<unknown> {
  return graphRequest(
    config,
    'DELETE',
    `/sites/${pathSegment(siteId)}/lists/${pathSegment(listId)}/items/${pathSegment(itemId)}`,
    { scopes: ['Sites.ReadWrite.All'] }
  );
}

export async function listSharePointDrives(config: AppConfig, siteId: string, top?: number): Promise<unknown> {
  return graphRequest(config, 'GET', `/sites/${pathSegment(siteId)}/drives`, {
    query: {
      $top: clampTop(top, 25, 100),
      $select: 'id,name,description,driveType,webUrl,createdDateTime,lastModifiedDateTime,owner,quota'
    },
    scopes: ['Files.Read.All']
  });
}

export async function listSharePointDriveItems(
  config: AppConfig,
  driveId: string,
  parentItemId?: string,
  top?: number
): Promise<unknown> {
  const path = parentItemId
    ? `/drives/${pathSegment(driveId)}/items/${pathSegment(parentItemId)}/children`
    : `/drives/${pathSegment(driveId)}/root/children`;
  return graphRequest(config, 'GET', path, {
    query: {
      $top: clampTop(top, 25, 100),
      $select: 'id,name,webUrl,size,createdDateTime,lastModifiedDateTime,folder,file,parentReference,sharepointIds'
    },
    scopes: ['Files.Read.All']
  });
}

export async function searchSharePointDriveItems(
  config: AppConfig,
  driveId: string,
  query: string,
  top?: number
): Promise<unknown> {
  const escaped = escapeODataString(query.trim());
  return graphRequest(config, 'GET', `/drives/${pathSegment(driveId)}/root/search(q='${escaped}')`, {
    query: {
      $top: clampTop(top, 10, 50),
      $select: 'id,name,webUrl,size,createdDateTime,lastModifiedDateTime,folder,file,parentReference,sharepointIds'
    },
    scopes: ['Files.Read.All']
  });
}

export async function getSharePointDriveItem(
  config: AppConfig,
  driveId: string,
  itemId: string
): Promise<unknown> {
  return graphRequest(config, 'GET', `/drives/${pathSegment(driveId)}/items/${pathSegment(itemId)}`, {
    query: {
      $select: 'id,name,webUrl,size,createdDateTime,lastModifiedDateTime,folder,file,parentReference,sharepointIds,listItem'
    },
    scopes: ['Files.Read.All']
  });
}

export async function createSharePointDriveFolder(
  config: AppConfig,
  driveId: string,
  name: string,
  parentItemId?: string,
  conflictBehavior: 'rename' | 'replace' | 'fail' = 'rename'
): Promise<unknown> {
  const path = parentItemId
    ? `/drives/${pathSegment(driveId)}/items/${pathSegment(parentItemId)}/children`
    : `/drives/${pathSegment(driveId)}/root/children`;
  return graphRequest(config, 'POST', path, {
    body: {
      name,
      folder: {},
      '@microsoft.graph.conflictBehavior': conflictBehavior
    },
    scopes: ['Files.ReadWrite.All']
  });
}

export async function uploadSmallSharePointFile(
  config: AppConfig,
  driveId: string,
  path: string,
  content: string,
  contentEncoding: 'text' | 'base64' = 'text',
  contentType?: string,
  conflictBehavior: 'replace' | 'rename' | 'fail' = 'replace'
): Promise<unknown> {
  const bytes = contentEncoding === 'base64' ? base64ToBytes(content) : content;
  return graphRequest(config, 'PUT', `/drives/${pathSegment(driveId)}/root:/${drivePath(path)}:/content`, {
    query: { '@microsoft.graph.conflictBehavior': conflictBehavior },
    body: bytes,
    contentType: contentType ?? (contentEncoding === 'text' ? 'text/plain; charset=utf-8' : 'application/octet-stream'),
    scopes: ['Files.ReadWrite.All']
  });
}

export async function deleteSharePointDriveItem(
  config: AppConfig,
  driveId: string,
  itemId: string
): Promise<unknown> {
  return graphRequest(config, 'DELETE', `/drives/${pathSegment(driveId)}/items/${pathSegment(itemId)}`, {
    scopes: ['Files.ReadWrite.All']
  });
}

export async function listMyJoinedTeams(config: AppConfig): Promise<unknown> {
  return graphRequest(config, 'GET', '/me/joinedTeams', {
    scopes: ['Team.ReadBasic.All']
  });
}

export async function listTeamChannels(config: AppConfig, teamId: string): Promise<unknown> {
  return graphRequest(config, 'GET', `/teams/${pathSegment(teamId)}/channels`, {
    query: {
      $select: 'id,displayName,description,webUrl,membershipType,isFavoriteByDefault'
    },
    scopes: ['Channel.ReadBasic.All']
  });
}

export async function listChannelMessages(
  config: AppConfig,
  teamId: string,
  channelId: string,
  top?: number
): Promise<unknown> {
  return graphRequest(config, 'GET', `/teams/${pathSegment(teamId)}/channels/${pathSegment(channelId)}/messages`, {
    query: {
      $top: clampTop(top, 10, 50)
    },
    scopes: ['ChannelMessage.Read.All']
  });
}

export async function listChannelMessageReplies(
  config: AppConfig,
  teamId: string,
  channelId: string,
  messageId: string,
  top?: number
): Promise<unknown> {
  return graphRequest(
    config,
    'GET',
    `/teams/${pathSegment(teamId)}/channels/${pathSegment(channelId)}/messages/${pathSegment(messageId)}/replies`,
    {
      query: { $top: clampTop(top, 10, 50) },
      scopes: ['ChannelMessage.Read.All']
    }
  );
}

export async function sendChannelMessage(
  config: AppConfig,
  teamId: string,
  channelId: string,
  content: string,
  contentIsHtml?: boolean
): Promise<unknown> {
  return graphRequest(config, 'POST', `/teams/${pathSegment(teamId)}/channels/${pathSegment(channelId)}/messages`, {
    body: {
      body: teamsBody(content, contentIsHtml)
    },
    scopes: ['ChannelMessage.Send']
  });
}

export async function replyToChannelMessage(
  config: AppConfig,
  teamId: string,
  channelId: string,
  messageId: string,
  content: string,
  contentIsHtml?: boolean
): Promise<unknown> {
  return graphRequest(
    config,
    'POST',
    `/teams/${pathSegment(teamId)}/channels/${pathSegment(channelId)}/messages/${pathSegment(messageId)}/replies`,
    {
      body: {
        body: teamsBody(content, contentIsHtml)
      },
      scopes: ['ChannelMessage.Send']
    }
  );
}

export async function listMyChats(config: AppConfig, top?: number): Promise<unknown> {
  return graphRequest(config, 'GET', '/me/chats', {
    query: {
      $top: clampTop(top, 20, 50),
      $select: 'id,topic,chatType,createdDateTime,lastUpdatedDateTime,webUrl'
    },
    scopes: ['Chat.Read']
  });
}

export async function listChatMessages(config: AppConfig, chatId: string, top?: number): Promise<unknown> {
  return graphRequest(config, 'GET', `/chats/${pathSegment(chatId)}/messages`, {
    query: {
      $top: clampTop(top, 10, 50)
    },
    scopes: ['Chat.Read']
  });
}

export async function sendChatMessage(
  config: AppConfig,
  chatId: string,
  content: string,
  contentIsHtml?: boolean
): Promise<unknown> {
  return graphRequest(config, 'POST', `/chats/${pathSegment(chatId)}/messages`, {
    body: {
      body: teamsBody(content, contentIsHtml)
    },
    scopes: ['ChatMessage.Send']
  });
}

export async function listMyContacts(config: AppConfig, top?: number): Promise<unknown> {
  return graphRequest(config, 'GET', '/me/contacts', {
    query: {
      $top: clampTop(top, 20, 100),
      $select: 'id,displayName,givenName,surname,emailAddresses,businessPhones,mobilePhone,companyName,jobTitle'
    },
    scopes: ['Contacts.Read']
  });
}

export async function createContact(
  config: AppConfig,
  input: {
    givenName?: string;
    surname?: string;
    displayName?: string;
    emailAddresses?: ContactEmailInput[];
    businessPhones?: string[];
    mobilePhone?: string;
    companyName?: string;
    jobTitle?: string;
  }
): Promise<unknown> {
  return graphRequest(config, 'POST', '/me/contacts', {
    body: input,
    scopes: ['Contacts.ReadWrite']
  });
}

export async function updateContact(
  config: AppConfig,
  contactId: string,
  input: {
    givenName?: string;
    surname?: string;
    displayName?: string;
    emailAddresses?: ContactEmailInput[];
    businessPhones?: string[];
    mobilePhone?: string;
    companyName?: string;
    jobTitle?: string;
  }
): Promise<unknown> {
  return graphRequest(config, 'PATCH', `/me/contacts/${pathSegment(contactId)}`, {
    body: input,
    scopes: ['Contacts.ReadWrite']
  });
}

export async function deleteContact(config: AppConfig, contactId: string): Promise<unknown> {
  return graphRequest(config, 'DELETE', `/me/contacts/${pathSegment(contactId)}`, {
    scopes: ['Contacts.ReadWrite']
  });
}
