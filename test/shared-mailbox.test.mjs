import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';

import { loadConfig } from '../dist/config.js';
import { listCalendarOverview, listMailOverview, readMailMessage, searchMailOverview } from '../dist/graph-scenarios.js';
import { permissionReport } from '../dist/permissions-report.js';
import { runWithRequestContext } from '../dist/request-context.js';

const graph = 'https://microsoftgraph.chinacloudapi.cn/v1.0';

async function graphTestConfig(t, overrides = {}) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'ms365-21v-shared-'));
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
    disabledGraphScopes: [],
    allowedGraphScopes: [],
    graphResilience: {
      maxRetries: 0,
      initialBackoffMs: 0,
      backoffMultiplier: 1,
      circuitBreakerThreshold: 100,
      circuitBreakerCooldownMs: 1000,
      timeoutMs: 1000
    },
    ...overrides
  };
}

const json = body => new Response(JSON.stringify(body), { status: 200, headers: { 'content-type': 'application/json' } });

async function withMockFetch(handler, callback) {
  const original = globalThis.fetch;
  const requests = [];
  globalThis.fetch = async (url, init = {}) => {
    const request = { url: new URL(String(url)), init };
    requests.push(request);
    return handler(request);
  };
  try {
    const result = await runWithRequestContext(
      { requestId: 'shared-test', inboundClaims: { oid: 'me-id', preferred_username: 'Me@Example.cn' } },
      callback
    );
    return { result, requests };
  } finally {
    globalThis.fetch = original;
  }
}

const isScopeError = scope => error => error.code === 'scope_not_enabled' && error.message.includes(scope);

test('mail tools read a shared mailbox through /users/{mailbox} and keep my own mailbox on /me', async (t) => {
  const config = await graphTestConfig(t);
  const page = json({
    value: [{ id: 'm1', subject: 'Order', receivedDateTime: '2026-11-03T01:00:00Z' }],
    '@odata.nextLink': `${graph}/users/team%40example.cn/mailFolders/inbox/messages?%24skip=1`
  });

  const shared = await withMockFetch(() => page, () => listMailOverview(config, { mailbox: 'team@example.cn' }));
  assert.equal(shared.requests[0].url.pathname, '/v1.0/users/team%40example.cn/mailFolders/inbox/messages');
  assert.equal(shared.result.mailbox, 'team@example.cn');

  const next = await withMockFetch(() => json({ value: [] }), () => listMailOverview(config, { cursor: shared.result.nextCursor }));
  assert.equal(next.requests[0].url.pathname, '/v1.0/users/team%40example.cn/mailFolders/inbox/messages');
  assert.equal(next.result.mailbox, 'team@example.cn');

  const own = await withMockFetch(() => json({ value: [] }), () => listMailOverview(config, { mailbox: 'ME@example.cn' }));
  assert.equal(own.requests[0].url.pathname, '/v1.0/me/mailFolders/inbox/messages');
  assert.equal(own.result.mailbox, undefined);

  const search = await withMockFetch(
    () => json({ value: [] }),
    () => searchMailOverview(config, { query: 'invoice', mailbox: 'team@example.cn' })
  );
  assert.equal(search.requests[0].url.pathname, '/v1.0/users/team%40example.cn/messages');
  assert.equal(search.requests[0].url.searchParams.get('$search'), '"invoice"');

  const detail = await withMockFetch(
    ({ url }) =>
      url.pathname.endsWith('/attachments')
        ? json({ value: [{ id: 'a1', name: 'po.pdf', size: 10 }] })
        : json({ id: 'm1', hasAttachments: true, body: { contentType: 'text', content: 'Hello' } }),
    () => readMailMessage(config, { messageId: 'm1', mailbox: 'team@example.cn' })
  );
  assert.deepEqual(
    detail.requests.map(request => request.url.pathname),
    ['/v1.0/users/team%40example.cn/messages/m1', '/v1.0/users/team%40example.cn/messages/m1/attachments']
  );
  assert.equal(detail.result.mailbox, 'team@example.cn');
  assert.equal(detail.result.attachments[0].name, 'po.pdf');

  const blocked = { ...config, disabledGraphScopes: ['Mail.Read.Shared'] };
  await assert.rejects(() => listMailOverview(blocked, { mailbox: 'team@example.cn' }), isScopeError('Mail.Read.Shared'));
  await assert.rejects(() => listMailOverview(blocked, { cursor: shared.result.nextCursor }), isScopeError('Mail.Read.Shared'));
  await assert.rejects(
    () => listMailOverview(config, { cursor: `${graph}/users/team%40example.cn/events?$skip=1` }),
    /different resource/
  );
});

test('calendar view reads a shared calendar, follows its cursor and needs Calendars.Read.Shared', async (t) => {
  const config = await graphTestConfig(t);
  const cursor = `${graph}/users/boss%40example.cn/calendarView?startDateTime=2026-11-03T00:00:00Z&endDateTime=2026-11-10T00:00:00Z&%24skip=50`;
  const first = await withMockFetch(
    () =>
      json({
        value: [{ id: 'e1', subject: 'Board', start: { dateTime: '2026-11-04T09:00:00' }, end: { dateTime: '2026-11-04T10:00:00' } }],
        '@odata.nextLink': cursor
      }),
    () =>
      listCalendarOverview(config, {
        startDateTime: '2026-11-03T00:00:00Z',
        endDateTime: '2026-11-10T00:00:00Z',
        calendarOwner: 'boss@example.cn'
      })
  );
  assert.equal(first.requests[0].url.pathname, '/v1.0/users/boss%40example.cn/calendarView');
  assert.equal(first.result.calendarOwner, 'boss@example.cn');
  assert.equal(first.result.nextCursor, cursor);

  const next = await withMockFetch(
    () => json({ value: [] }),
    () => listCalendarOverview(config, { startDateTime: 'x', endDateTime: 'y', cursor })
  );
  assert.equal(next.requests[0].url.href, cursor);
  assert.equal(next.requests[0].init.headers.Prefer, 'outlook.timezone="China Standard Time"');
  assert.equal(next.result.calendarOwner, 'boss@example.cn');

  await assert.rejects(
    () =>
      listCalendarOverview(
        { ...config, allowedGraphScopes: ['Calendars.Read'] },
        { startDateTime: 'x', endDateTime: 'y', calendarOwner: 'boss@example.cn' }
      ),
    isScopeError('Calendars.Read.Shared')
  );
  await assert.rejects(
    () => listCalendarOverview(config, { startDateTime: 'x', endDateTime: 'y', cursor: `${graph}/me/messages?$skip=1` }),
    /different resource/
  );
});

test('a shared mailbox or calendar that cannot be read says why', async (t) => {
  const config = await graphTestConfig(t);
  const notFound = code =>
    new Response(JSON.stringify({ error: { code, message: 'not found' } }), { status: 404, headers: { 'content-type': 'application/json' } });
  const range = { startDateTime: '2026-11-04T00:00:00Z', endDateTime: '2026-11-05T00:00:00Z' };

  await assert.rejects(
    () => withMockFetch(() => notFound('ErrorItemNotFound'), () => listCalendarOverview(config, { ...range, calendarOwner: 'boss@example.cn' })),
    error => error.code === 'not_shared' && /not shared with you/.test(error.message) && /calendar_get_schedule/.test(error.message)
  );
  await assert.rejects(
    () => withMockFetch(() => notFound('ErrorInvalidUser'), () => listMailOverview(config, { mailbox: 'someone@other.example' })),
    error => error.code === 'unknown_user' && /users_search/.test(error.message)
  );
  await assert.rejects(
    () => withMockFetch(() => notFound('ErrorItemNotFound'), () => searchMailOverview(config, { query: 'x', mailbox: 'team@example.cn' })),
    error => error.code === 'not_shared' && /mailbox of team@example.cn/.test(error.message)
  );
  await assert.rejects(
    () => withMockFetch(() => notFound('ErrorItemNotFound'), () => listMailOverview(config, { folder: 'missing-folder' })),
    error => error.name === 'GraphHttpError' && error.status === 404,
    'my own mailbox keeps the Graph error'
  );
});

test('the permission report lists the shared mailbox and calendar scopes as optional', async (t) => {
  const report = permissionReport(await graphTestConfig(t, { allowedGraphScopes: ['Mail.Read', 'Calendars.Read'] }));
  const optional = Object.fromEntries(report.optionalScopes.map(item => [item.scope, item]));
  assert.equal(optional['Mail.Read.Shared'].enabled, false);
  assert.deepEqual(
    [...optional['Mail.Read.Shared'].tools].sort(),
    ['mail_get_attachment', 'mail_get_message', 'mail_list_messages', 'mail_search_messages']
  );
  assert.deepEqual(optional['Calendars.Read.Shared'].tools, ['calendar_list_view']);
});
