import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';

import { defaultDirectTools, loadConfig } from '../dist/config.js';
import {
  cursorPath,
  getFileDownloadUrl,
  listCalendarOverview,
  listChatMessagesWindow,
  listChatsOverview,
  listMailOverview,
  mailListFilter,
  readMailMessage,
  searchFiles,
  sharingToken,
  toUtcInstant
} from '../dist/graph-scenarios.js';
import { createGraphMcpServer } from '../dist/mcp-server.js';
import { runWithRequestContext } from '../dist/request-context.js';
import { jsonResult } from '../dist/tools/results.js';
import { htmlToText, pruneGraphPayload } from '../dist/text.js';

const graph = 'https://microsoftgraph.chinacloudapi.cn/v1.0';
const me = { oid: 'me-id', preferred_username: 'Me@Example.cn' };

async function graphTestConfig(t) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'ms365-21v-scenario-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const tokenCachePath = path.join(directory, 'token.json');
  await fs.writeFile(
    tokenCachePath,
    JSON.stringify({
      accessToken: 'test-graph-token',
      tokenType: 'Bearer',
      expiresAt: Math.floor(Date.now() / 1000) + 3600,
      acquiredAt: Math.floor(Date.now() / 1000)
    }),
    { mode: 0o600 }
  );
  return {
    ...loadConfig(),
    tokenCachePath,
    deviceCodeCachePath: `${tokenCachePath}.device-code`,
    graphBaseUrl: graph,
    graphResource: 'https://microsoftgraph.chinacloudapi.cn',
    auditLogEnabled: false,
    graphResilience: {
      maxRetries: 0,
      initialBackoffMs: 0,
      backoffMultiplier: 1,
      circuitBreakerThreshold: 100,
      circuitBreakerCooldownMs: 1000,
      timeoutMs: 1000
    }
  };
}

function json(body, status = 200) {
  return new Response(JSON.stringify(body), { status, headers: { 'content-type': 'application/json' } });
}

async function withMockFetch(handler, callback) {
  const original = globalThis.fetch;
  const requests = [];
  globalThis.fetch = async (url, init = {}) => {
    const request = { url: new URL(String(url)), init };
    requests.push(request);
    return handler(request);
  };
  try {
    const result = await runWithRequestContext({ requestId: 'scenario-test', inboundClaims: me }, callback);
    return { result, requests };
  } finally {
    globalThis.fetch = original;
  }
}

test('time inputs default to China Standard Time and reject invalid values', () => {
  assert.equal(toUtcInstant('2026-10-12', 'since'), '2026-10-11T16:00:00Z');
  assert.equal(toUtcInstant('2026-10-12T09:30', 'since'), '2026-10-12T01:30:00Z');
  assert.equal(toUtcInstant('2026-10-12T09:30:00Z', 'since'), '2026-10-12T09:30:00Z');
  assert.equal(toUtcInstant('2026-10-12T09:30:00+02:00', 'since'), '2026-10-12T07:30:00Z');
  assert.equal(toUtcInstant(undefined, 'since'), undefined);
  assert.throws(() => toUtcInstant('yesterday', 'since'), /since must be an ISO 8601/);
});

test('HTML from Outlook and Teams becomes readable text', () => {
  const text = htmlToText(
    '<p>Hi&nbsp;<at id="0">Wang</at>,</p><table><tr><th>Item</th><th>Owner</th></tr><tr><td>Plan</td><td>Li</td></tr></table>' +
      '<ul><li>One</li><li>Two &amp; three</li></ul><emoji alt="👍"></emoji><img src="x"><br>Bye<attachment id="1"></attachment>'
  );
  assert.equal(text, 'Hi Wang,\n\nItem | Owner\nPlan | Li\n\n- One\n- Two & three\n👍[image]\nBye');
});

test('tool results drop Graph metadata and nulls and are not indented', () => {
  assert.deepEqual(pruneGraphPayload({ '@odata.context': 'x', '@odata.nextLink': 'n', a: null, b: [{ c: null, d: 1 }] }), {
    '@odata.nextLink': 'n',
    b: [{ d: 1 }]
  });
  const result = jsonResult({ value: [{ id: '1', topic: null }] });
  assert.equal(result.content[0].text, '{"value":[{"id":"1"}]}');
});

test('cursors are limited to the same Graph resource', async (t) => {
  const config = await graphTestConfig(t);
  assert.equal(
    cursorPath(config, `${graph}/me/chats?$skiptoken=abc`, /^\/(me|users\/[^/]+)\/chats$/),
    `${graph}/me/chats?$skiptoken=abc`
  );
  assert.throws(() => cursorPath(config, 'https://evil.example.cn/v1.0/me/chats', /^\/me\/chats$/), /nextCursor/);
  assert.throws(() => cursorPath(config, `${graph}/me/messages?$skip=10`, /^\/me\/chats$/), /different resource/);
  assert.throws(() => cursorPath(config, 'not a url', /^\/me\/chats$/), /nextCursor/);
});

const chatPage = {
  '@odata.nextLink': `${graph}/me/chats?$skiptoken=page2`,
  value: [
    {
      id: 'chat-new',
      chatType: 'oneOnOne',
      topic: null,
      webUrl: 'https://teams/chat-new',
      members: [
        { userId: 'me-id', displayName: 'Me', email: 'me@example.cn' },
        { userId: 'zhang', displayName: 'Zhang Wei', email: 'wei.zhang@example.cn' }
      ],
      lastMessagePreview: {
        createdDateTime: '2026-10-12T02:00:00Z',
        messageType: 'message',
        body: { contentType: 'html', content: '<p>Please send the <b>report</b></p>' },
        from: { user: { id: 'zhang', displayName: 'Zhang Wei' } }
      }
    },
    {
      id: 'chat-group',
      chatType: 'group',
      topic: 'Line 1 PoC',
      members: [{ userId: 'me-id', displayName: 'Me' }, { userId: 'li', displayName: 'Li Na' }],
      lastMessagePreview: {
        createdDateTime: '2026-10-11T15:00:00Z',
        messageType: 'message',
        body: { contentType: 'text', content: 'ok' },
        from: { user: { id: 'me-id', displayName: 'Me' } }
      }
    },
    {
      id: 'chat-old',
      chatType: 'group',
      topic: 'Old',
      members: [],
      lastMessagePreview: { createdDateTime: '2026-10-01T00:00:00Z', messageType: 'message', body: { content: 'x' } }
    }
  ]
};

test('chat list expands members and last message, filters by recency and person, and hides me', async (t) => {
  const config = await graphTestConfig(t);
  const { result, requests } = await withMockFetch(
    () => json(chatPage),
    () => listChatsOverview(config, { since: '2026-10-12T00:00', top: 10 })
  );
  assert.equal(requests.length, 1);
  const query = requests[0].url.searchParams;
  assert.equal(requests[0].url.pathname, '/v1.0/me/chats');
  assert.equal(query.get('$expand'), 'members,lastMessagePreview');
  assert.equal(query.get('$orderby'), 'lastMessagePreview/createdDateTime desc');
  assert.equal(query.get('$top'), '50');
  assert.deepEqual(result.chats.map(chat => chat.id), ['chat-new']);
  assert.equal(result.nextCursor, undefined);
  const [chat] = result.chats;
  assert.equal(chat.title, 'Zhang Wei');
  assert.deepEqual(chat.members, ['Zhang Wei']);
  assert.deepEqual(chat.lastMessage, {
    at: '2026-10-12T02:00:00Z',
    from: 'Zhang Wei',
    fromMe: undefined,
    preview: 'Please send the report',
    type: undefined
  });

  const byPerson = await withMockFetch(
    ({ url }) => (url.searchParams.get('$skiptoken') ? json({ value: [] }) : json(chatPage)),
    () => listChatsOverview(config, { query: 'li na' })
  );
  assert.equal(byPerson.requests.length, 2);
  assert.deepEqual(byPerson.result.chats.map(chat => chat.title), ['Line 1 PoC']);

  const repeated = await withMockFetch(
    () => json(chatPage),
    () => listChatsOverview(config, { query: 'li na' })
  );
  assert.equal(repeated.requests.length, 5);
  assert.deepEqual(repeated.result.chats.map(chat => chat.id), ['chat-group']);
  assert.equal(byPerson.result.chats[0].lastMessage.fromMe, true);
});

test('chat list without filters returns one page and its cursor', async (t) => {
  const config = await graphTestConfig(t);
  const { result, requests } = await withMockFetch(
    () => json({ ...chatPage, value: chatPage.value.slice(0, 2) }),
    () => listChatsOverview(config, { top: 2 })
  );
  assert.equal(requests[0].url.searchParams.get('$top'), '2');
  assert.equal(result.count, 2);
  assert.equal(result.nextCursor, `${graph}/me/chats?$skiptoken=page2`);
});

const messagePage = {
  '@odata.nextLink': `${graph}/chats/chat-new/messages?$skiptoken=older`,
  value: [
    {
      id: 'm3',
      messageType: 'message',
      createdDateTime: '2026-10-12T03:00:00Z',
      lastModifiedDateTime: '2026-10-12T03:00:00Z',
      from: { user: { id: 'zhang', displayName: 'Zhang Wei' } },
      body: { contentType: 'html', content: '<p><at id="0">Me</at> can you review?</p>' },
      mentions: [{ mentionText: 'Me', mentioned: { user: { id: 'me-id' } } }],
      attachments: [{ contentType: 'reference', name: 'plan.docx', contentUrl: 'https://tenant.sharepoint.cn/plan.docx' }]
    },
    {
      id: 'm2',
      messageType: 'systemEventMessage',
      createdDateTime: '2026-10-12T02:30:00Z',
      body: { content: '<systemEventMessage/>' }
    },
    {
      id: 'm1',
      messageType: 'message',
      createdDateTime: '2026-10-12T02:00:00Z',
      from: { user: { id: 'me-id', displayName: 'Me' } },
      body: { contentType: 'text', content: 'Draft attached' }
    },
    {
      id: 'm0',
      messageType: 'message',
      createdDateTime: '2026-10-11T08:00:00Z',
      lastModifiedDateTime: '2026-10-12T02:10:00Z',
      from: { user: { id: 'zhang', displayName: 'Zhang Wei' } },
      body: { contentType: 'text', content: 'edited old message' }
    }
  ]
};

test('chat messages use the 21V-supported filters and return a readable time window', async (t) => {
  const config = await graphTestConfig(t);
  const { result, requests } = await withMockFetch(
    () => json(messagePage),
    () => listChatMessagesWindow(config, { chatId: 'chat-new', since: '2026-10-12T09:00' })
  );
  const query = requests[0].url.searchParams;
  assert.equal(query.get('$orderby'), 'lastModifiedDateTime desc');
  assert.equal(query.get('$filter'), 'lastModifiedDateTime gt 2026-10-12T01:00:00Z');
  assert.deepEqual(result.messages.map(message => message.id), ['m1', 'm3']);
  assert.equal(result.messages[0].fromMe, true);
  assert.equal(result.messages[1].text, 'Me can you review?');
  assert.equal(result.messages[1].mentionsMe, true);
  assert.deepEqual(result.messages[1].attachments, [
    { type: 'file', name: 'plan.docx', url: 'https://tenant.sharepoint.cn/plan.docx' }
  ]);
  assert.equal(result.nextCursor, `${graph}/chats/chat-new/messages?$skiptoken=older`);

  const untilOnly = await withMockFetch(
    () => json({ value: [] }),
    () => listChatMessagesWindow(config, { chatId: 'chat-new', until: '2026-10-12' })
  );
  assert.equal(untilOnly.requests[0].url.searchParams.get('$orderby'), 'createdDateTime desc');
  assert.equal(untilOnly.requests[0].url.searchParams.get('$filter'), 'createdDateTime lt 2026-10-11T16:00:00Z');

  const system = await withMockFetch(
    () => json(messagePage),
    () => listChatMessagesWindow(config, { chatId: 'chat-new', includeSystemMessages: true })
  );
  assert.deepEqual(system.result.messages.map(message => message.id), ['m0', 'm1', 'm2', 'm3']);
  assert.equal(system.result.messages[2].type, 'systemEventMessage');
});

test('mail filters lead with receivedDateTime as Graph requires', () => {
  assert.equal(mailListFilter({}), undefined);
  assert.equal(
    mailListFilter({ flaggedOnly: true, unreadOnly: true }),
    "receivedDateTime ge 1900-01-01T00:00:00Z and flag/flagStatus eq 'flagged' and isRead eq false"
  );
  assert.equal(
    mailListFilter({ since: '2026-10-11T16:00:00Z', until: '2026-10-12T16:00:00Z', from: "O'Brien@Example.com", importance: 'high' }),
    "receivedDateTime ge 2026-10-11T16:00:00Z and receivedDateTime lt 2026-10-12T16:00:00Z and importance eq 'high' and from/emailAddress/address eq 'o''brien@example.com'"
  );
});

test('mail list reads the inbox by default and marks my recipient role', async (t) => {
  const config = await graphTestConfig(t);
  const { result, requests } = await withMockFetch(
    () =>
      json({
        value: [
          {
            id: 'mail-1',
            subject: 'Review',
            from: { emailAddress: { name: 'Zhang Wei', address: 'wei.zhang@example.cn' } },
            toRecipients: [{ emailAddress: { name: 'Someone', address: 'someone@example.cn' } }],
            ccRecipients: [{ emailAddress: { name: 'Me', address: 'me@example.cn' } }],
            receivedDateTime: '2026-10-12T01:00:00Z',
            bodyPreview: 'Please review\r\nthe plan',
            isRead: false,
            importance: 'high',
            flag: { flagStatus: 'flagged' },
            '@odata.etag': 'W/"1"'
          }
        ]
      }),
    () => listMailOverview(config, { since: '2026-10-12', flaggedOnly: true })
  );
  assert.equal(requests[0].url.pathname, '/v1.0/me/mailFolders/inbox/messages');
  assert.equal(
    requests[0].url.searchParams.get('$filter'),
    "receivedDateTime ge 2026-10-11T16:00:00Z and flag/flagStatus eq 'flagged'"
  );
  const [row] = result.messages;
  assert.equal(row.from, 'Zhang Wei <wei.zhang@example.cn>');
  assert.equal(row.myRole, 'cc');
  assert.equal(row.preview, 'Please review the plan');
  assert.equal(row.flagged, true);
  assert.equal(row.importance, 'high');

  const all = await withMockFetch(() => json({ value: [] }), () => listMailOverview(config, { folder: 'all' }));
  assert.equal(all.requests[0].url.pathname, '/v1.0/me/messages');
  assert.equal(all.requests[0].url.searchParams.get('$filter'), null);

  await assert.rejects(() => listMailOverview(config, { from: 'Zhang Wei' }), /mail_search_messages/);
});

test('mail details return plain text in pages with the attachment list', async (t) => {
  const config = await graphTestConfig(t);
  const body = 'A'.repeat(250) + '\n\n\n\nEnd';
  const { result, requests } = await withMockFetch(
    ({ url }) =>
      url.pathname.endsWith('/attachments')
        ? json({ value: [{ id: 'att-1', name: 'plan.pdf', contentType: 'application/pdf', size: 1200, isInline: false }] })
        : json({
            id: 'mail-1',
            subject: 'Plan',
            hasAttachments: true,
            uniqueBody: { contentType: 'text', content: body }
          }),
    () => readMailMessage(config, { messageId: 'mail-1', onlyNewContent: true, maxChars: 200 })
  );
  assert.equal(requests[0].init.headers.Prefer, 'outlook.body-content-type="text"');
  assert.match(requests[0].url.searchParams.get('$select'), /uniqueBody$/);
  assert.equal(result.body.text.length, 200);
  assert.equal(result.body.totalChars, 255);
  assert.equal(result.body.nextOffset, 200);
  assert.deepEqual(result.attachments, [
    { id: 'att-1', name: 'plan.pdf', contentType: 'application/pdf', size: 1200, inline: undefined }
  ]);

  const rest = await withMockFetch(
    ({ url }) => (url.pathname.endsWith('/attachments') ? json({ value: [] }) : json({ id: 'mail-1', body: { contentType: 'text', content: body } })),
    () => readMailMessage(config, { messageId: 'mail-1', maxChars: 200, offset: 200, includeAttachments: false })
  );
  assert.equal(rest.requests.length, 1);
  assert.equal(rest.result.body.text, 'A'.repeat(50) + '\n\nEnd');
  assert.equal(rest.result.body.nextOffset, undefined);
});

test('calendar view returns compact events', async (t) => {
  const config = await graphTestConfig(t);
  const { result, requests } = await withMockFetch(
    () =>
      json({
        value: [
          {
            id: 'event-1',
            subject: 'Weekly',
            start: { dateTime: '2026-10-12T09:00:00.0000000', timeZone: 'China Standard Time' },
            end: { dateTime: '2026-10-12T10:00:00.0000000', timeZone: 'China Standard Time' },
            organizer: { emailAddress: { name: 'Li Na', address: 'na.li@example.cn' } },
            attendees: [{ emailAddress: { name: 'Me', address: 'me@example.cn' } }],
            responseStatus: { response: 'accepted' },
            seriesMasterId: 'series',
            isCancelled: false
          }
        ]
      }),
    () => listCalendarOverview(config, { startDateTime: '2026-10-11T16:00:00Z', endDateTime: '2026-10-12T16:00:00Z' })
  );
  assert.equal(requests[0].init.headers.Prefer, 'outlook.timezone="China Standard Time"');
  assert.deepEqual(result.events[0], {
    id: 'event-1',
    subject: 'Weekly',
    start: '2026-10-12T09:00',
    end: '2026-10-12T10:00',
    isAllDay: undefined,
    isCancelled: undefined,
    location: undefined,
    organizer: 'Li Na <na.li@example.cn>',
    myResponse: 'accepted',
    showAs: undefined,
    isOnlineMeeting: undefined,
    recurring: true,
    attendeeCount: 1,
    attendees: ['Me'],
    preview: undefined,
    webLink: undefined
  });
});

const driveItem = {
  id: 'item-1',
  name: 'Plan.docx',
  webUrl: 'https://tenant.sharepoint.cn/sites/Project/Shared%20Documents/Plan.docx',
  size: 2048,
  lastModifiedDateTime: '2026-10-10T08:00:00Z',
  lastModifiedBy: { user: { displayName: 'Li Na' } },
  parentReference: { driveId: 'drive-1', path: '/drives/drive-1/root:/Specs' },
  file: { mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' },
  '@microsoft.graph.downloadUrl': 'https://tenant.sharepoint.cn/download?token=short'
};

test('file search uses Microsoft Search and falls back to library search', async (t) => {
  const config = await graphTestConfig(t);
  const { result, requests } = await withMockFetch(
    () => json({ value: [{ hitsContainers: [{ hits: [{ summary: 'the <c0>plan</c0> draft<ddd/>', resource: driveItem }] }] }] }),
    () => searchFiles(config, { query: 'plan', fileTypes: ['docx'], modifiedAfter: '2026-10-01' })
  );
  assert.equal(requests[0].url.pathname, '/v1.0/search/query');
  assert.equal(
    JSON.parse(requests[0].init.body).requests[0].query.queryString,
    'plan (filetype:docx) LastModifiedTime>=2026-09-30'
  );
  assert.equal(result.engine, 'microsoft_search');
  assert.deepEqual(result.files[0], {
    name: 'Plan.docx',
    fileType: 'docx',
    isFolder: undefined,
    location: 'SharePoint: Project',
    path: '/Specs',
    size: 2048,
    lastModifiedAt: '2026-10-10T08:00:00Z',
    lastModifiedBy: 'Li Na',
    summary: 'the plan draft…',
    webUrl: driveItem.webUrl,
    driveId: 'drive-1',
    itemId: 'item-1'
  });

  const fallback = await withMockFetch(
    ({ url }) =>
      url.pathname === '/v1.0/search/query'
        ? json({ error: { code: 'BadRequest' } }, 400)
        : json({ value: [{ ...driveItem, webUrl: 'https://tenant-my.sharepoint.cn/personal/me/Plan.docx' }] }),
    () => searchFiles(config, { query: "O'Neil plan" })
  );
  assert.equal(fallback.result.engine, 'library_search');
  assert.equal(fallback.requests[1].url.pathname, "/v1.0/me/drive/root/search(q='O''Neil%20plan')");
  assert.equal(fallback.result.files[0].location, 'OneDrive');
  assert.match(fallback.result.notes.join(' '), /HTTP 400/);
});

test('download URLs resolve sharing links, convert to PDF and inline small text files', async (t) => {
  const config = await graphTestConfig(t);
  const link = 'https://tenant.sharepoint.cn/:w:/s/Project/abc?e=1';
  const shared = await withMockFetch(
    () => json(driveItem),
    () => getFileDownloadUrl(config, { url: link })
  );
  assert.equal(shared.requests[0].url.pathname, `/v1.0/shares/${sharingToken(link)}/driveItem`);
  assert.equal(shared.requests[0].init.headers.Prefer, 'redeemSharingLinkIfNecessary');
  assert.equal(shared.result.downloadUrl, driveItem['@microsoft.graph.downloadUrl']);
  assert.equal(shared.result.textContent, undefined);
  assert.match(shared.result.parsingHint, /format pdf/);
  assert.ok(sharingToken(link).startsWith('u!'));
  assert.doesNotMatch(sharingToken(link), /[+/=]/);

  const pdf = await withMockFetch(
    ({ url }) =>
      url.pathname.endsWith('/content')
        ? new Response(null, { status: 302, headers: { location: 'https://tenant.sharepoint.cn/pdf?token=short' } })
        : json(driveItem),
    () => getFileDownloadUrl(config, { driveId: 'drive-1', itemId: 'item-1', format: 'pdf' })
  );
  assert.equal(pdf.requests[1].url.searchParams.get('format'), 'pdf');
  assert.equal(pdf.requests[1].init.redirect, 'manual');
  assert.equal(pdf.result.downloadUrl, 'https://tenant.sharepoint.cn/pdf?token=short');
  assert.equal(pdf.result.mimeType, 'application/pdf');

  const notes = await withMockFetch(
    ({ url }) =>
      url.hostname === 'tenant.sharepoint.cn'
        ? new Response('# Notes\r\n\r\n\r\nline', { status: 200 })
        : json({ ...driveItem, name: 'notes.md', size: 20, file: { mimeType: 'text/markdown' } }),
    () => getFileDownloadUrl(config, { itemId: 'item-1' })
  );
  assert.equal(notes.requests[0].url.pathname, '/v1.0/me/drive/items/item-1');
  assert.equal(notes.requests[1].init.headers, undefined);
  assert.equal(notes.result.textContent, '# Notes\n\nline');

  await assert.rejects(
    () => withMockFetch(() => json({ ...driveItem, folder: { childCount: 2 } }), () => getFileDownloadUrl(config, { itemId: 'item-1' })),
    /is a folder/
  );
  await assert.rejects(
    () => withMockFetch(() => json({ error: {} }, 404), () => getFileDownloadUrl(config, { url: link })),
    /Could not open this link \(HTTP 404\)/
  );
});

test('hybrid is the default and lists the scenario tools plus gateway controls', () => {
  const saved = { mode: process.env.MCP_TOOL_EXPOSURE_MODE, direct: process.env.MCP_DIRECT_TOOLS };
  try {
    delete process.env.MCP_TOOL_EXPOSURE_MODE;
    delete process.env.MCP_DIRECT_TOOLS;
    const config = loadConfig();
    assert.equal(config.toolExposureMode, 'hybrid');
    assert.deepEqual(config.directTools, defaultDirectTools);
    process.env.MCP_DIRECT_TOOLS = 'none';
    assert.deepEqual(loadConfig().directTools, []);

    const server = createGraphMcpServer(
      {
        ...config,
        transport: 'http',
        inboundAuthDisabled: false,
        roleBasedFiltering: true,
        auditLogEnabled: false,
        disabledGraphScopes: [],
        sendMode: 'confirm'
      },
      { userRoles: ['mcp.admin'] }
    );
    assert.deepEqual(Object.keys(server._registeredTools).sort(), [
      'auth_status',
      'calendar_list_view',
      'confirm_execute',
      'files_get_download_url',
      'files_search',
      'gateway_execute_tool',
      'gateway_get_tool_schema',
      'gateway_search_tools',
      'mail_get_message',
      'mail_list_messages',
      'mail_search_messages',
      'teams_list_chat_messages',
      'teams_list_chats'
    ]);
  } finally {
    for (const [key, value] of [
      ['MCP_TOOL_EXPOSURE_MODE', saved.mode],
      ['MCP_DIRECT_TOOLS', saved.direct]
    ]) {
      if (value === undefined) delete process.env[key];
      else process.env[key] = value;
    }
  }
});

test('discovery search returns schemas for top hits and invalid arguments return the schema', async () => {
  const server = createGraphMcpServer(
    {
      ...loadConfig(),
      transport: 'http',
      inboundAuthDisabled: false,
      roleBasedFiltering: true,
      auditLogEnabled: false,
      disabledGraphScopes: [],
      toolExposureMode: 'hybrid',
      directTools: []
    },
    { userRoles: ['mcp.admin'] }
  );
  const call = (name, args) =>
    runWithRequestContext({ requestId: 'discovery', inboundClaims: { sub: 'u', roles: ['mcp.admin'] } }, () =>
      server._registeredTools[name].handler(args, {})
    );
  const search = await call('gateway_search_tools', { query: '接受会议邀请', limit: 5, schemaCount: 2 });
  const [first, second, third] = search.structuredContent.tools;
  assert.equal(first.name, 'calendar_respond_to_event');
  assert.equal(first.inputSchema.type, 'object');
  assert.ok(second.inputSchema);
  assert.equal(third.inputSchema, undefined);

  const free = await call('gateway_search_tools', { query: '查看谁有空', limit: 1, schemaCount: 0 });
  assert.equal(free.structuredContent.tools[0].name, 'calendar_get_schedule');

  const invalid = await call('gateway_execute_tool', { toolName: 'calendar_respond_to_event', parameters: {} });
  assert.equal(invalid.isError, true);
  assert.equal(invalid.structuredContent.error.code, 'invalid_tool_arguments');
  assert.equal(invalid.structuredContent.error.inputSchema.type, 'object');
});

const oneDriveItem = {
  ...driveItem,
  id: 'item-2',
  name: 'Notes.docx',
  webUrl: 'https://tenant-my.sharepoint.cn/personal/me/Documents/Notes.docx',
  parentReference: { driveId: 'drive-me', driveType: 'business' },
  '@microsoft.graph.downloadUrl': 'https://tenant-my.sharepoint.cn/download?token=short'
};

test('without mcp.sharepoint, file tools keep to OneDrive and refuse SharePoint files', async (t) => {
  const config = await graphTestConfig(t);
  const search = await withMockFetch(
    () => json({ value: [{ hitsContainers: [{ hits: [{ resource: driveItem }, { resource: oneDriveItem }] }] }] }),
    () => searchFiles(config, { query: 'plan', top: 5, includeSharePoint: false })
  );
  assert.equal(JSON.parse(search.requests[0].init.body).requests[0].size, 50);
  assert.deepEqual(search.result.files.map(file => file.location), ['OneDrive']);
  assert.match(search.result.notes.join(' '), /mcp\.sharepoint/);

  const isPermissionDenied = error => error.code === 'permission_denied' && /mcp\.sharepoint/.test(error.message);
  await assert.rejects(
    () => searchFiles(config, { query: 'plan', site: 'https://tenant.sharepoint.cn/sites/Project', includeSharePoint: false }),
    isPermissionDenied
  );
  await assert.rejects(() => searchFiles(config, { query: 'plan', scope: 'sharepoint', includeSharePoint: false }), isPermissionDenied);

  await assert.rejects(
    () =>
      withMockFetch(
        () => json({ ...driveItem, parentReference: { driveId: 'drive-1', driveType: 'documentLibrary' } }),
        () => getFileDownloadUrl(config, { driveId: 'drive-1', itemId: 'item-1', includeSharePoint: false })
      ),
    isPermissionDenied
  );
  await assert.rejects(
    () => withMockFetch(() => json(driveItem), () => getFileDownloadUrl(config, { url: 'https://tenant.sharepoint.cn/:w:/s/Project/abc', includeSharePoint: false })),
    isPermissionDenied
  );
  const own = await withMockFetch(
    () => json(oneDriveItem),
    () => getFileDownloadUrl(config, { itemId: 'item-2', includeSharePoint: false })
  );
  assert.equal(own.result.location, 'OneDrive');
  assert.equal(own.result.downloadUrl, oneDriveItem['@microsoft.graph.downloadUrl']);
});

test('file tools read mcp.sharepoint from the caller App Roles', async (t) => {
  const config = await graphTestConfig(t);
  const server = createGraphMcpServer(
    {
      ...config,
      transport: 'http',
      inboundAuthDisabled: false,
      roleBasedFiltering: true,
      disabledGraphScopes: [],
      toolExposureMode: 'hybrid',
      directTools: ['files_search'],
      directToolCategories: []
    },
    { userRoles: ['mcp.drive'] }
  );
  const searchAs = async roles => {
    const original = globalThis.fetch;
    globalThis.fetch = async () =>
      json({ value: [{ hitsContainers: [{ hits: [{ resource: driveItem }, { resource: oneDriveItem }] }] }] });
    try {
      return await runWithRequestContext({ requestId: 'file-roles', inboundClaims: { ...me, roles } }, () =>
        server._registeredTools.files_search.handler({ query: 'plan' }, {})
      );
    } finally {
      globalThis.fetch = original;
    }
  };
  const locations = result => result.structuredContent.files.map(file => file.location);
  assert.deepEqual(locations(await searchAs(['mcp.drive'])), ['OneDrive']);
  assert.deepEqual(locations(await searchAs(['mcp.drive', 'mcp.sharepoint'])), ['SharePoint: Project', 'OneDrive']);
  assert.deepEqual(locations(await searchAs(['mcp.admin'])), ['SharePoint: Project', 'OneDrive']);
  const denied = await searchAs(['mcp.sharepoint']);
  assert.equal(denied.isError, true);
  assert.equal(denied.structuredContent.error.code, 'permission_denied');
});
