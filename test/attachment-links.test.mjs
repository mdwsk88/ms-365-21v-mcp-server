import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import http from 'node:http';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';

import express from 'express';

import { AttachmentLinkStore, attachmentLinkStore, mountAttachmentRoutes } from '../dist/attachment-links.js';
import { loadConfig } from '../dist/config.js';
import { getMailAttachment } from '../dist/graph-scenarios.js';
import { runWithRequestContext } from '../dist/request-context.js';

const graph = 'https://microsoftgraph.chinacloudapi.cn/v1.0';

async function testConfig(t, overrides = {}) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'ms365-21v-attachments-'));
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
    publicPathPrefix: '',
    auditLogEnabled: false,
    disabledGraphScopes: [],
    allowedGraphScopes: [],
    attachmentUrlTtlSeconds: 240,
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
      { requestId: 'attachment-test', inboundClaims: { oid: 'user-1' } },
      callback
    );
    return { result, requests };
  } finally {
    globalThis.fetch = original;
  }
}

const pdfMeta = {
  '@odata.type': '#microsoft.graph.fileAttachment',
  id: 'att-1',
  name: '季度报告 Q3.pdf',
  contentType: 'application/pdf',
  size: 73_000,
  isInline: false
};

test('link store hands a ticket out once, expires it and caps links per user', () => {
  let now = 1_000_000;
  const store = new AttachmentLinkStore(() => now);
  const link = { graphPath: '/me/x/$value', accessToken: 'token', name: 'a.pdf', contentType: 'application/pdf', userKey: 'u1' };
  const first = store.create(link, 240);
  assert.match(first.ticket, /^[A-Za-z0-9_-]{43}$/);
  assert.equal(store.redeem(first.ticket)?.name, 'a.pdf');
  assert.equal(store.redeem(first.ticket), undefined);

  const late = store.create(link, 240);
  now += 241_000;
  assert.equal(store.redeem(late.ticket), undefined);

  const tickets = Array.from({ length: 25 }, () => store.create(link, 240).ticket);
  assert.equal(store.size, 20);
  assert.equal(store.redeem(tickets[0]), undefined);
  assert.ok(store.redeem(tickets[24]));
});

test('mail attachments return text inline, links as source URLs and files as single-use URLs', async (t) => {
  const config = await testConfig(t);
  attachmentLinkStore.clear();
  t.after(() => attachmentLinkStore.clear());

  const textFile = await withMockFetch(
    ({ url }) =>
      url.pathname.endsWith('/$value')
        ? new Response('名称,数量\r\nA,1', { status: 200 })
        : json({ ...pdfMeta, name: 'data.csv', contentType: 'text/csv', size: 20 }),
    () => getMailAttachment(config, { messageId: 'm/1', attachmentId: 'att-1' })
  );
  assert.equal(textFile.requests[0].url.pathname, '/v1.0/me/messages/m%2F1/attachments/att-1');
  assert.equal(textFile.requests[0].url.searchParams.get('$select'), 'id,name,contentType,size,isInline,lastModifiedDateTime');
  assert.equal(textFile.result.textContent, '名称,数量\nA,1');
  assert.equal(textFile.result.downloadUrl, undefined);

  const reference = await withMockFetch(
    ({ url }) =>
      url.searchParams.get('$select')
        ? json({ '@odata.type': '#microsoft.graph.referenceAttachment', id: 'att-2', name: 'Plan' })
        : json({ '@odata.type': '#microsoft.graph.referenceAttachment', id: 'att-2', sourceUrl: 'https://tenant.sharepoint.cn/plan' }),
    () => getMailAttachment(config, { messageId: 'm1', attachmentId: 'att-2' })
  );
  assert.equal(reference.result.kind, 'link');
  assert.equal(reference.result.sourceUrl, 'https://tenant.sharepoint.cn/plan');

  const pdf = await withMockFetch(
    () => json(pdfMeta),
    () => getMailAttachment(config, { messageId: 'm1', attachmentId: 'att-1' })
  );
  assert.equal(pdf.requests.length, 1);
  assert.match(pdf.result.downloadUrl, /^https:\/\/mcp\.example\.cn\/attachments\/[A-Za-z0-9_-]{43}$/);
  assert.match(pdf.result.downloadUrlNotice, /Single-use/);
  assert.equal(attachmentLinkStore.size, 1);

  const item = await withMockFetch(
    () => json({ ...pdfMeta, '@odata.type': '#microsoft.graph.itemAttachment', name: 'Re: Budget', contentType: undefined }),
    () => getMailAttachment(config, { messageId: 'm1', attachmentId: 'att-3' })
  );
  assert.equal(item.result.name, 'Re: Budget.eml');
  assert.equal(item.result.contentType, 'message/rfc822');

  const stdio = await withMockFetch(
    ({ url }) =>
      url.pathname.endsWith('/$value')
        ? new Response(Buffer.from('%PDF-1.7'), { status: 200, headers: { 'content-type': 'application/pdf' } })
        : json(pdfMeta),
    () => getMailAttachment({ ...config, transport: 'stdio' }, { messageId: 'm1', attachmentId: 'att-1' })
  );
  assert.equal(Buffer.from(stdio.result.contentBase64, 'base64').toString(), '%PDF-1.7');

  await assert.rejects(
    () =>
      getMailAttachment(
        { ...config, disabledGraphScopes: ['Mail.Read.Shared'] },
        { messageId: 'm1', attachmentId: 'att-1', mailbox: 'team@example.cn' }
      ),
    error => error.code === 'scope_not_enabled'
  );
});

test('the download route streams the attachment once with the minting user token', async (t) => {
  const config = await testConfig(t);
  attachmentLinkStore.clear();
  t.after(() => attachmentLinkStore.clear());
  const minted = await withMockFetch(
    () => json(pdfMeta),
    () => getMailAttachment(config, { messageId: 'm1', attachmentId: 'att-1' })
  );
  const pathname = new URL(minted.result.downloadUrl).pathname;

  const app = express();
  mountAttachmentRoutes(app, config);
  const server = app.listen(0, '127.0.0.1');
  await new Promise(resolve => server.once('listening', resolve));
  t.after(() => server.close());
  const port = server.address().port;
  const get = () =>
    new Promise((resolve, reject) => {
      http
        .get({ host: '127.0.0.1', port, path: pathname }, response => {
          const chunks = [];
          response.on('data', chunk => chunks.push(chunk));
          response.on('end', () => resolve({ status: response.statusCode, headers: response.headers, body: Buffer.concat(chunks) }));
        })
        .on('error', reject);
    });

  const original = globalThis.fetch;
  const graphRequests = [];
  globalThis.fetch = async (url, init = {}) => {
    graphRequests.push({ url: String(url), init });
    return new Response(Buffer.from('%PDF-1.7 bytes'), {
      status: 200,
      headers: { 'content-type': 'application/pdf', 'content-length': '14' }
    });
  };
  t.after(() => {
    globalThis.fetch = original;
  });

  const first = await get();
  assert.equal(first.status, 200);
  assert.equal(first.body.toString(), '%PDF-1.7 bytes');
  assert.equal(first.headers['content-type'], 'application/pdf');
  assert.equal(first.headers['cache-control'], 'no-store');
  assert.equal(first.headers['x-content-type-options'], 'nosniff');
  assert.match(first.headers['content-disposition'], /filename\*=UTF-8''%E5%AD%A3%E5%BA%A6/);
  assert.equal(graphRequests.length, 1);
  assert.equal(graphRequests[0].url, `${graph}/me/messages/m1/attachments/att-1/$value`);
  assert.equal(graphRequests[0].init.headers.Authorization, 'Bearer test-graph-token');

  const second = await get();
  assert.equal(second.status, 410);
  assert.match(second.body.toString(), /link_expired/);
  assert.equal(graphRequests.length, 1);
});
