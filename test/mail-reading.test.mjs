import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';

import { loadConfig } from '../dist/config.js';
import { getMailAttachment, readMailMessage, readMailMessages, shapeMailRow } from '../dist/graph-scenarios.js';
import { createGraphMcpServer } from '../dist/mcp-server.js';
import { runWithRequestContext } from '../dist/request-context.js';
import { cleanMailText, cleanUrl, normalizeText } from '../dist/text.js';

const graph = 'https://microsoftgraph.chinacloudapi.cn/v1.0';

async function graphTestConfig(t) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'ms365-21v-mail-'));
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
    transport: 'http',
    tokenCachePath,
    deviceCodeCachePath: `${tokenCachePath}.device-code`,
    graphBaseUrl: graph,
    graphResource: 'https://microsoftgraph.chinacloudapi.cn',
    publicBaseUrl: 'https://mcp.example.cn',
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
    const result = await runWithRequestContext(
      { requestId: 'mail-test', inboundClaims: { oid: 'me-id', preferred_username: 'me@example.cn' } },
      callback
    );
    return { result, requests };
  } finally {
    globalThis.fetch = original;
  }
}

const safeLink = target =>
  `https://nam12.safelinks.protection.outlook.com/?url=${encodeURIComponent(target)}&data=05%7C02%7Cme%40example.cn%7Cabc%7C0%7C638000000000000000&sdata=c2lnbmF0dXJl%3D&reserved=0`;

test('links lose Safe Links wrappers and tracking parameters but nothing else', () => {
  assert.equal(cleanUrl(safeLink('https://example.com/page?id=1&utm_source=news')), 'https://example.com/page?id=1');
  assert.equal(
    cleanUrl(
      'https://can01.safelinks.protection.partner.outlook.cn/?url=https%3A%2F%2Fcontoso.sharepoint.cn%2Fsites%2FAI&data=x&reserved=0'
    ),
    'https://contoso.sharepoint.cn/sites/AI'
  );
  assert.equal(cleanUrl(safeLink(safeLink('https://example.com/a'))), 'https://example.com/a');
  assert.equal(cleanUrl('https://example.com/a?b=%20c&fbclid=1&mc_eid=2'), 'https://example.com/a?b=%20c');
  assert.equal(cleanUrl('https://Example.com'), 'https://Example.com', 'untouched links keep their spelling');
  assert.equal(cleanUrl(safeLink('javascript:alert(1)')).startsWith('https://nam12.safelinks'), true);
});

test('mail text keeps one readable copy of each link and drops padding', () => {
  const tracker = `https://click.example.com/ls/click?upn=${'x'.repeat(600)}`;
  const sharePoint = `https://contoso.sharepoint.cn/:w:/r/sites/AI/Doc.aspx?sourcedoc=${'%E5%AD%A3'.repeat(60)}`;
  const body = [
    'Hi\u00a0team,\u200b',
    '\u034f \u200c \u034f \u200c \u034f \u200c',
    `Please read https://example.com/plan<${safeLink('https://example.com/plan')}> first.`,
    `Agenda<${safeLink('https://example.com/agenda?utm_campaign=q4')}>`,
    `Newsletter<${tracker}> and doc<${sharePoint}>`,
    `(see https://example.com/a?utm_source=x).`,
    '详见https://example.com/b?fbclid=1，谢谢',
    'zhang.san@example.cn<mailto:zhang.san@example.cn> or Li Si<mailto:li.si@example.cn?subject=Hi>',
    'a+b@example.cn<mailto:a+b@example.cn>',
    'Family 👨\u200d👩\u200d👧',
    '[cid:image001.png@01DB1234.56789AB0]'
  ].join('\n');

  const { text, linksOmitted } = cleanMailText(normalizeText(body));
  assert.equal(linksOmitted, 1);
  assert.deepEqual(text.split('\n'), [
    'Hi team,',
    '',
    'Please read https://example.com/plan first.',
    'Agenda<https://example.com/agenda>',
    `Newsletter [link: click.example.com] and doc<${sharePoint}>`,
    '(see https://example.com/a).',
    '详见https://example.com/b，谢谢',
    'zhang.san@example.cn or Li Si <li.si@example.cn>',
    'a+b@example.cn',
    'Family 👨\u200d👩\u200d👧',
    '[image]'
  ]);
});

test('mail lists clean previews', () => {
  const row = shapeMailRow(
    { id: 'm1', bodyPreview: `\u034f \u200c Read ${safeLink('https://example.com/x?utm_medium=mail')} now` },
    { email: 'me@example.cn' }
  );
  assert.equal(row.preview, 'Read https://example.com/x now');
});

test('mail details are cleaned unless clean is false', async (t) => {
  const config = await graphTestConfig(t);
  const body = `Join<${safeLink('https://example.com/join')}>\u200b`;
  const respond = () => json({ id: 'mail-1', subject: 'Invite', body: { contentType: 'text', content: body } });

  const cleaned = await withMockFetch(respond, () => readMailMessage(config, { messageId: 'mail-1' }));
  assert.equal(cleaned.result.body.text, 'Join<https://example.com/join>');
  assert.equal(cleaned.result.body.linksOmitted, undefined);

  const raw = await withMockFetch(respond, () => readMailMessage(config, { messageId: 'mail-1', clean: false }));
  assert.match(raw.result.body.text, /safelinks\.protection\.outlook\.com/);

  const html = await withMockFetch(
    () => json({ id: 'mail-1', body: { contentType: 'html', content: '<p>Hi</p>' } }),
    () => readMailMessage(config, { messageId: 'mail-1', bodyFormat: 'html' })
  );
  assert.equal(html.result.body.text, '<p>Hi</p>');
});

test('several messages are read in one call, in order, with per-message errors', async (t) => {
  const config = await graphTestConfig(t);
  let active = 0;
  let peak = 0;
  const { result, requests } = await withMockFetch(
    async ({ url }) => {
      active += 1;
      peak = Math.max(peak, active);
      await new Promise(resolve => setTimeout(resolve, 5));
      active -= 1;
      const id = decodeURIComponent(url.pathname.split('/').pop());
      if (id === 'missing') {
        return json({ error: { code: 'ErrorItemNotFound', message: 'The specified object was not found in the store.' } }, 404);
      }
      return json({ id, subject: `Subject ${id}`, body: { contentType: 'text', content: `Body ${id} `.repeat(2000) } });
    },
    () =>
      readMailMessages(config, {
        messageIds: ['m1', 'missing', 'm2', 'm1', 'm3', 'm4', 'm5', 'm6'],
        includeAttachments: false
      })
  );
  assert.equal(result.count, 7);
  assert.equal(result.failed, 1);
  assert.deepEqual(
    result.messages.map(message => message.id),
    ['m1', 'missing', 'm2', 'm3', 'm4', 'm5', 'm6']
  );
  assert.deepEqual(result.messages[1].error, {
    status: 404,
    code: 'ErrorItemNotFound',
    message: 'The specified object was not found in the store.'
  });
  assert.equal(result.messages[0].body.text.length, 8000);
  assert.ok(result.messages[0].body.nextOffset);
  assert.equal(requests.length, 7);
  assert.ok(peak <= 4, `at most four Graph calls run at once (saw ${peak})`);

  await assert.rejects(() => readMailMessages(config, { messageIds: ['m1'], offset: 10 }), /single messageId/);
  await assert.rejects(
    () => readMailMessages(config, { messageIds: Array.from({ length: 11 }, (_, index) => `m${index}`) }),
    /at most 10/
  );
});

test('mail_get_message accepts messageId or messageIds', async (t) => {
  const config = { ...(await graphTestConfig(t)), toolExposureMode: 'direct', roleBasedFiltering: false };
  const server = createGraphMcpServer(config);
  const handler = server._registeredTools.mail_get_message.handler;
  const call = input =>
    withMockFetch(
      ({ url }) => json({ id: decodeURIComponent(url.pathname.split('/').pop()), body: { contentType: 'text', content: 'Hi' } }),
      () => handler(input, {})
    );

  const batch = await call({ messageId: 'a', messageIds: ['b'], includeAttachments: false });
  assert.deepEqual(
    batch.result.structuredContent.messages.map(message => message.id),
    ['a', 'b']
  );
  const single = await call({ messageId: 'a' });
  assert.equal(single.result.structuredContent.body.text, 'Hi');
  const missing = await call({});
  assert.equal(missing.result.isError, true);
  assert.match(missing.result.structuredContent.error.message, /messageId/);
});

test('text attachments keep tabs and indentation', async (t) => {
  const config = await graphTestConfig(t);
  const content = '\ufeffname\tcount\r\nA\t1\n\nconfig:\n  nested:\n    value: 1\n';
  const { result } = await withMockFetch(
    ({ url }) =>
      url.pathname.endsWith('/$value')
        ? new Response(content, { status: 200 })
        : json({ '@odata.type': '#microsoft.graph.fileAttachment', id: 'att-1', name: 'data.tsv', size: 60 }),
    () => getMailAttachment(config, { messageId: 'm1', attachmentId: 'att-1' })
  );
  assert.equal(result.textContent, 'name\tcount\nA\t1\n\nconfig:\n  nested:\n    value: 1\n');
});
