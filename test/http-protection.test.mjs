import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import fs from 'node:fs/promises';
import { createServer as createNetServer } from 'node:net';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';

import express from 'express';

import { ConfigError, loadConfig } from '../dist/config.js';
import { mountOAuthBridgeRoutes, redirectUriProblem } from '../dist/oauth-bridge.js';
import { RateLimiter } from '../dist/rate-limit.js';

function unsignedJwt(payload) {
  const encode = value => Buffer.from(JSON.stringify(value)).toString('base64url');
  return `${encode({ alg: 'none', typ: 'JWT' })}.${encode(payload)}.`;
}

async function availablePort() {
  const server = createNetServer();
  await new Promise((resolve, reject) => {
    server.once('error', reject);
    server.listen(0, '127.0.0.1', resolve);
  });
  const { port } = server.address();
  await new Promise(resolve => server.close(resolve));
  return port;
}

const env = (values, callback) => {
  const previous = Object.fromEntries(Object.keys(values).map(key => [key, process.env[key]]));
  Object.assign(process.env, values);
  try {
    return callback();
  } finally {
    for (const [key, value] of Object.entries(previous)) {
      if (value === undefined) delete process.env[key];
      else process.env[key] = value;
    }
  }
};

test('the rate limiter counts a fixed window per key', () => {
  let now = 0;
  const limiter = new RateLimiter(2, 60_000, () => now);
  assert.equal(limiter.take('a').allowed, true);
  assert.equal(limiter.take('a').allowed, true);
  assert.deepEqual(limiter.take('a'), { allowed: false, retryAfterSeconds: 60 });
  assert.equal(limiter.take('b').allowed, true);
  now = 45_000;
  assert.deepEqual(limiter.take('a'), { allowed: false, retryAfterSeconds: 15 });
  now = 60_000;
  assert.equal(limiter.take('a').allowed, true);
});

test('client redirect URIs: loopback and native apps always, other hosts only over https when listed', () => {
  const none = { oauthRedirectHosts: [] };
  for (const allowed of [
    'http://127.0.0.1:5555/callback',
    'http://localhost:33418/callback',
    'http://[::1]:8080/cb',
    'cursor://anysphere.cursor-retrieval/oauth/user-ms365/callback',
    'com.example.app:/oauth2redirect'
  ]) {
    assert.equal(redirectUriProblem(none, allowed), undefined, allowed);
  }
  assert.match(redirectUriProblem(none, 'https://claude.ai/api/mcp/auth_callback'), /not in MCP_OAUTH_REDIRECT_HOSTS/);
  assert.match(redirectUriProblem(none, 'javascript:alert(1)'), /javascript/);
  assert.match(redirectUriProblem(none, 'mailto:someone@example.com'), /mailto/);
  assert.match(redirectUriProblem(none, 'http://127.0.0.1/cb#fragment'), /fragment/);
  assert.match(redirectUriProblem(none, '/relative/callback'), /absolute/);

  const listed = { oauthRedirectHosts: ['claude.ai', '*.example.cn'] };
  assert.equal(redirectUriProblem(listed, 'https://claude.ai/api/mcp/auth_callback'), undefined);
  assert.equal(redirectUriProblem(listed, 'https://agent.example.cn/callback'), undefined);
  assert.match(redirectUriProblem(listed, 'https://example.cn/callback'), /not in/);
  assert.match(redirectUriProblem(listed, 'http://claude.ai/callback'), /must use https/);
  assert.equal(redirectUriProblem({ oauthRedirectHosts: ['*'] }, 'https://anything.example.org/cb'), undefined);
});

test('trust proxy, redirect hosts and limits are parsed from the environment', () => {
  const keys = {
    MCP_HTTP_HOST: '127.0.0.1',
    MCP_TRUST_PROXY: '',
    MCP_OAUTH_REDIRECT_HOSTS: '',
    MCP_RATE_LIMIT_ENABLED: '',
    MCP_HTTP_MAX_SESSIONS: ''
  };
  env(keys, () => {
    const config = loadConfig();
    assert.equal(config.trustProxy, 'loopback');
    assert.deepEqual(config.oauthRedirectHosts, []);
    assert.equal(config.rateLimitEnabled, true);
    assert.equal(config.httpMaxSessions, 200);
    assert.equal(config.oauthDynamicRegistration, true);
  });
  env({ ...keys, MCP_HTTP_HOST: '0.0.0.0' }, () => assert.equal(loadConfig().trustProxy, false));
  env({ ...keys, MCP_TRUST_PROXY: '2' }, () => assert.equal(loadConfig().trustProxy, 2));
  env({ ...keys, MCP_TRUST_PROXY: 'loopback, 10.0.0.0/8' }, () => assert.equal(loadConfig().trustProxy, 'loopback, 10.0.0.0/8'));
  env({ ...keys, MCP_TRUST_PROXY: 'true' }, () => assert.throws(() => loadConfig(), ConfigError));
  env({ ...keys, MCP_OAUTH_REDIRECT_HOSTS: 'https://Claude.ai/api, *.Example.cn,chatgpt.com' }, () =>
    assert.deepEqual(loadConfig().oauthRedirectHosts, ['claude.ai', '*.example.cn', 'chatgpt.com'])
  );
});

async function bridgeApp(t, overrides, state) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'mcp-bridge-policy-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const statePath = path.join(directory, 'state.json');
  if (state) await fs.writeFile(statePath, JSON.stringify(state));
  const config = {
    ...loadConfig(),
    tenantId: 'tenant-id',
    authorityHost: 'https://login.partner.microsoftonline.cn',
    oauthBridgeEnabled: true,
    oauthBridgeMicrosoftClientType: 'confidential_web',
    oauthBridgeMicrosoftClientId: 'web-client-id',
    oauthBridgeMicrosoftClientSecret: 'web-client-secret',
    oauthBridgeIssuer: 'https://mcp.example.cn',
    oauthBridgeRedirectUri: 'https://mcp.example.cn/oauth/microsoft/callback',
    oauthBridgeStatePath: statePath,
    publicPathPrefix: '',
    oauthDynamicRegistration: true,
    oauthRedirectHosts: [],
    ...overrides
  };
  const app = express();
  app.use(express.json());
  mountOAuthBridgeRoutes(app, config);
  const server = app.listen(0, '127.0.0.1');
  await new Promise(resolve => server.once('listening', resolve));
  t.after(() => server.close());
  return { origin: `http://127.0.0.1:${server.address().port}`, statePath };
}

const register = (origin, redirectUris) =>
  fetch(`${origin}/oauth/register`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ client_name: 'Test client', redirect_uris: redirectUris })
  });

const authorize = (origin, clientId, redirectUri) => {
  const url = new URL(`${origin}/oauth/authorize`);
  url.searchParams.set('response_type', 'code');
  url.searchParams.set('client_id', clientId);
  url.searchParams.set('redirect_uri', redirectUri);
  url.searchParams.set('code_challenge', 'x'.repeat(43));
  url.searchParams.set('code_challenge_method', 'S256');
  return fetch(url, { redirect: 'manual' });
};

test('registration refuses redirect URIs outside the policy, and authorize re-checks older clients', async (t) => {
  const now = Math.floor(Date.now() / 1000);
  const legacy = {
    clientId: 'legacy-client',
    redirectUris: ['https://attacker.example.com/callback'],
    grantTypes: ['authorization_code', 'refresh_token'],
    responseTypes: ['code'],
    tokenEndpointAuthMethod: 'none',
    createdAt: now - 100
  };
  const { origin } = await bridgeApp(t, {}, { clients: [legacy], sessions: [] });

  const refused = await register(origin, ['https://attacker.example.com/callback']);
  assert.equal(refused.status, 400);
  assert.equal((await refused.json()).error, 'invalid_redirect_uri');
  assert.equal((await register(origin, Array.from({ length: 11 }, (_, i) => `http://127.0.0.1:${4000 + i}/cb`))).status, 400);

  const accepted = await register(origin, ['http://127.0.0.1:5555/callback']);
  assert.equal(accepted.status, 201);
  const { client_id: clientId } = await accepted.json();
  const redirect = await authorize(origin, clientId, 'http://127.0.0.1:5555/callback');
  assert.equal(redirect.status, 302);
  assert.match(redirect.headers.get('location'), /^https:\/\/login\.partner\.microsoftonline\.cn\/tenant-id\/oauth2\/v2\.0\/authorize\?/);

  const old = await authorize(origin, 'legacy-client', 'https://attacker.example.com/callback');
  assert.equal(old.status, 400);
  assert.match((await old.json()).error_description, /MCP_OAUTH_REDIRECT_HOSTS/);
});

test('dynamic registration can be switched off and keeps at most 1000 clients', async (t) => {
  const off = await bridgeApp(t, { oauthDynamicRegistration: false });
  const metadata = await (await fetch(`${off.origin}/.well-known/oauth-authorization-server`)).json();
  assert.equal(metadata.registration_endpoint, undefined);
  assert.equal((await register(off.origin, ['http://127.0.0.1:5555/callback'])).status, 403);

  const now = Math.floor(Date.now() / 1000);
  const client = (id, age) => ({
    clientId: id,
    redirectUris: ['http://127.0.0.1:1/cb'],
    grantTypes: ['authorization_code'],
    responseTypes: ['code'],
    tokenEndpointAuthMethod: 'none',
    createdAt: now - age
  });
  const clients = Array.from({ length: 1000 }, (_, index) => client(`client-${index}`, 10_000 - index));
  const full = await bridgeApp(t, {}, {
    clients,
    sessions: [{ clientId: 'client-0', accessToken: 'a', refreshToken: 'r', scope: [], microsoftToken: {}, expiresAt: now + 600, createdAt: now }]
  });
  assert.equal((await register(full.origin, ['http://127.0.0.1:5555/callback'])).status, 201);
  const ids = JSON.parse(await fs.readFile(full.statePath, 'utf8')).clients.map(entry => entry.clientId);
  assert.equal(ids.length, 1000);
  assert.ok(ids.includes('client-0'), 'a client with a live session is kept');
  assert.ok(!ids.includes('client-1'), 'the least recently used idle client is dropped');
});

test('the HTTP server limits requests per address and per user and caps MCP sessions', async (t) => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'mcp-http-protection-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const statePath = path.join(directory, 'state.json');
  const now = Math.floor(Date.now() / 1000);
  const session = (user, accessToken) => ({
    clientId: 'test-client',
    accessToken,
    refreshToken: `${accessToken}-refresh`,
    scope: ['access_as_user'],
    microsoftToken: {
      accessToken: unsignedJwt({ sub: `${user}-sub`, oid: `${user}-oid`, roles: ['mcp.mail'], exp: now + 3600 }),
      tokenType: 'Bearer',
      expiresAt: now + 3600,
      acquiredAt: now
    },
    expiresAt: now + 3500,
    createdAt: now
  });
  await fs.writeFile(statePath, JSON.stringify({ clients: [], sessions: [session('alice', 'token-alice'), session('bob', 'token-bob')] }));

  const port = await availablePort();
  const origin = `http://127.0.0.1:${port}`;
  let output = '';
  const child = spawn(process.execPath, ['dist/index.js'], {
    env: {
      ...process.env,
      MCP_TRANSPORT: 'http',
      MCP_HTTP_HOST: '127.0.0.1',
      MCP_HTTP_PORT: String(port),
      MCP_HTTP_PATH: '/mcp',
      MCP_HTTP_STATELESS: 'false',
      MCP_OAUTH_HTTP_PATH: '',
      MCP_PUBLIC_BASE_URL: origin,
      MCP_RESOURCE_URL: '',
      MCP_RESOURCE_METADATA_URL: '',
      MCP_AUTHORIZATION_SERVERS: '',
      MCP_OAUTH_BRIDGE_ISSUER: '',
      MCP_OAUTH_BRIDGE_REDIRECT_URI: '',
      MCP_ALLOWED_HOSTS: '',
      MCP_ALLOW_UNAUTHENTICATED_DISCOVERY: 'false',
      MCP_INBOUND_AUTH_DISABLED: 'false',
      MCP_OAUTH_BRIDGE_ENABLED: 'true',
      MCP_OAUTH_BRIDGE_MICROSOFT_CLIENT_TYPE: 'confidential_web',
      MCP_OAUTH_BRIDGE_STATE_PATH: statePath,
      MCP_OAUTH_REDIRECT_HOSTS: '',
      MCP_OAUTH_DYNAMIC_REGISTRATION: '',
      MCP_TOKEN_AUDIENCE: 'api://api-client-id',
      MCP_AUTHORIZATION_SCOPES: 'api://api-client-id/access_as_user',
      MCP_REQUIRED_TOKEN_SCOPES: 'access_as_user',
      MCP_ROLE_BASED_FILTERING: 'false',
      MCP_AUDIT_LOG_ENABLED: 'false',
      MCP_TRUST_PROXY: '',
      MCP_RATE_LIMIT_ENABLED: 'true',
      MCP_RATE_LIMIT_USER_PER_MINUTE: '5',
      MCP_RATE_LIMIT_IP_PER_MINUTE: '20',
      MCP_HTTP_MAX_SESSIONS: '2',
      MCP_HTTP_SESSION_IDLE_MINUTES: '',
      MS_TENANT_ID: 'tenant-id',
      MS_CLIENT_ID: 'api-client-id',
      MS_CLIENT_SECRET: 'api-client-secret',
      MS_OAUTH_CLIENT_ID: '',
      MS_OAUTH_CLIENT_SECRET: ''
    },
    stdio: ['ignore', 'ignore', 'pipe']
  });
  child.stderr.on('data', chunk => {
    output += chunk.toString();
  });
  t.after(() => {
    if (child.exitCode === null) child.kill('SIGTERM');
  });
  const deadline = Date.now() + 10_000;
  for (;;) {
    if (child.exitCode !== null) throw new Error(`server exited early\n${output}`);
    if (Date.now() > deadline) throw new Error(`server did not start\n${output}`);
    if ((await fetch(`${origin}/healthz`).catch(() => undefined))?.ok) break;
    await new Promise(resolve => setTimeout(resolve, 50));
  }

  // Per address, from the address nginx forwards: OAuth registration gets a tenth of the limit.
  const fromAddress = (address, init = {}) => ({ ...init, headers: { ...init.headers, 'x-forwarded-for': address } });
  const registerFrom = address =>
    fetch(
      `${origin}/oauth/register`,
      fromAddress(address, {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ redirect_uris: ['http://127.0.0.1:5555/callback'] })
      })
    );
  assert.equal((await registerFrom('198.51.100.1')).status, 201);
  assert.equal((await registerFrom('198.51.100.1')).status, 201);
  const limited = await registerFrom('198.51.100.1');
  assert.equal(limited.status, 429);
  assert.ok(Number(limited.headers.get('retry-after')) > 0);
  assert.equal((await registerFrom('198.51.100.2')).status, 201, 'another address has its own budget');

  for (let index = 0; index < 20; index += 1) {
    assert.equal((await fetch(`${origin}/attachments/unknown`, fromAddress('203.0.113.7'))).status, 410);
  }
  assert.equal((await fetch(`${origin}/attachments/unknown`, fromAddress('203.0.113.7'))).status, 429);

  // Per user on /mcp.
  const mcp = (token, body, sessionId) =>
    fetch(`${origin}/mcp`, {
      method: 'POST',
      headers: {
        authorization: `Bearer ${token}`,
        'content-type': 'application/json',
        accept: 'application/json, text/event-stream',
        ...(sessionId ? { 'mcp-session-id': sessionId } : {})
      },
      body: JSON.stringify(body)
    });
  const ping = { jsonrpc: '2.0', id: 1, method: 'ping' };
  for (let index = 0; index < 5; index += 1) {
    const response = await mcp('token-alice', ping);
    assert.equal(response.status, 200, await response.text());
  }
  const overLimit = await mcp('token-alice', ping);
  assert.equal(overLimit.status, 429);
  assert.equal((await overLimit.json()).error.code, -32000);

  // Two sessions at most: the least recently used one is closed.
  const initialize = {
    jsonrpc: '2.0',
    id: 1,
    method: 'initialize',
    params: { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'test', version: '1' } }
  };
  const sessionIds = [];
  for (let index = 0; index < 3; index += 1) {
    const response = await mcp('token-bob', initialize);
    assert.equal(response.status, 200);
    await response.text();
    sessionIds.push(response.headers.get('mcp-session-id'));
  }
  // A POST with a closed session ID is still answered without a session; GET and DELETE need a live one.
  const end = sessionId =>
    fetch(`${origin}/mcp`, { method: 'DELETE', headers: { authorization: 'Bearer token-bob', 'mcp-session-id': sessionId } });
  assert.equal((await end(sessionIds[0])).status, 404);
  assert.equal((await end(sessionIds[2])).status, 200);
  assert.match(output, /"event":"evicted_oldest"/);
});
