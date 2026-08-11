import assert from 'node:assert/strict';
import test from 'node:test';

import { browserConfirmationStore } from '../dist/browser-confirmation.js';
import { mountConfirmationRoutes } from '../dist/confirmation-routes.js';
import { loadConfig } from '../dist/config.js';
import { jsonResult } from '../dist/tools/results.js';

function fakeApp() {
  const get = new Map();
  const post = new Map();
  return {
    get(path, handler) {
      get.set(path, handler);
    },
    post(path, handler) {
      post.set(path, handler);
    },
    routes: { get, post }
  };
}

function fakeResponse() {
  return {
    statusCode: 200,
    headers: {},
    body: undefined,
    set(name, value) {
      this.headers[name.toLowerCase()] = value;
      return this;
    },
    status(code) {
      this.statusCode = code;
      return this;
    },
    type(value) {
      this.headers['content-type'] = value;
      return this;
    },
    send(value) {
      this.body = value;
      return this;
    }
  };
}

function config() {
  return {
    ...loadConfig(),
    publicBaseUrl: 'https://mcp.example.cn/tools/MCP_21V',
    oauthBridgeIssuer: 'https://mcp.example.cn/tools/MCP_21V',
    publicPathPrefix: '/tools/MCP_21V',
    confirmOperations: ['send_email'],
    confirmTtlSeconds: 300
  };
}

function createPending() {
  return browserConfirmationStore.create(
    {
      name: 'mail_send',
      title: 'Send Mail',
      category: 'mail',
      module: 'Mail',
      description: '',
      isWriteOperation: true,
      requiresConfirmation: true,
      operationType: 'send_email',
      graphScopes: ['Mail.Send']
    },
    {
      subject: 'Quarterly <script>alert(1)</script>',
      body: 'Hi Fred,\n\nPlease review <img src=x onerror=alert(1)>.\n\nRegards,\nDawei',
      to: [{ name: 'Fred Zhong', email: 'fred.zhong@example.cn' }],
      cc: [{ email: 'reviewer@example.cn' }],
      bodyIsHtml: false,
      saveToSentItems: true
    },
    'user-a',
    300,
    async () => jsonResult({ sent: true }),
    config().publicBaseUrl
  );
}

test('web fallback renders a concrete mail preview without raw JSON or executable HTML', async (t) => {
  browserConfirmationStore.clear();
  t.after(() => browserConfirmationStore.clear());
  const app = fakeApp();
  mountConfirmationRoutes(app, config());
  const pending = createPending();

  const getHandler = app.routes.get.get('/tools/MCP_21V/confirm/:confirmToken');
  assert.ok(getHandler);
  const response = fakeResponse();
  getHandler({ params: { confirmToken: pending.confirmToken }, headers: {} }, response);

  assert.equal(response.statusCode, 200);
  assert.match(response.body, /发送邮件 \/ Send Mail/);
  assert.match(response.body, /收件人 \/ To/);
  assert.match(response.body, /密送 \/ Bcc/);
  assert.match(response.body, /Fred Zhong &lt;fred\.zhong@example\.cn&gt;/);
  assert.match(response.body, /Quarterly &lt;script&gt;alert\(1\)&lt;\/script&gt;/);
  assert.match(response.body, /Hi Fred,/);
  assert.match(response.body, /&lt;img src=x onerror=alert\(1\)&gt;/);
  assert.doesNotMatch(response.body, /<script>alert\(1\)<\/script>/);
  assert.doesNotMatch(response.body, /"toolName"/);
  assert.match(response.body, /确认并发送/);
  assert.match(response.body, /保存到已发送 \/ Save to Sent Items/);
  assert.match(response.headers['set-cookie'], /HttpOnly/);
  assert.match(response.headers['set-cookie'], /SameSite=Strict/);
  assert.match(response.headers['set-cookie'], /Secure/);
  assert.match(response.headers['content-security-policy'], /default-src 'none'/);

  const premature = await browserConfirmationStore.execute(pending.confirmToken, 'user-a');
  assert.equal(premature.structuredContent.error.code, 'human_confirmation_required');

  const approveHandler = app.routes.post.get('/tools/MCP_21V/confirm/:confirmToken/approve');
  const approveResponse = fakeResponse();
  approveHandler(
    {
      params: { confirmToken: pending.confirmToken },
      headers: { cookie: response.headers['set-cookie'].split(';')[0] }
    },
    approveResponse
  );
  assert.equal(approveResponse.statusCode, 200);
  assert.match(approveResponse.body, /操作已批准/);

  const executed = await browserConfirmationStore.execute(pending.confirmToken, 'user-a');
  assert.equal(executed.structuredContent.sent, true);
});

test('confirmation POST requires its same-site browser cookie', (t) => {
  browserConfirmationStore.clear();
  t.after(() => browserConfirmationStore.clear());
  const app = fakeApp();
  mountConfirmationRoutes(app, config());
  const pending = createPending();
  const response = fakeResponse();
  app.routes.post.get('/tools/MCP_21V/confirm/:confirmToken/approve')(
    { params: { confirmToken: pending.confirmToken }, headers: {} },
    response
  );
  assert.equal(response.statusCode, 403);
  assert.match(response.body, /无法批准/);
});
