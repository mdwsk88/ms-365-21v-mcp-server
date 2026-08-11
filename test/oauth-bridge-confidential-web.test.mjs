import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { spawn } from 'node:child_process';
import { createServer as createHttpServer } from 'node:http';
import { unlink } from 'node:fs/promises';
import { createServer as createNetServer } from 'node:net';
import os from 'node:os';
import path from 'node:path';
import { after, before, test } from 'node:test';

let child;
let microsoftServer;
let origin;
let microsoftOrigin;
let stderr = '';
let statePath;
const microsoftTokenRequests = [];

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
  const address = server.address();
  await new Promise(resolve => server.close(resolve));
  assert(address && typeof address === 'object');
  return address.port;
}

async function waitForServer(url) {
  const deadline = Date.now() + 10_000;
  while (Date.now() < deadline) {
    if (child.exitCode !== null) throw new Error(`server exited before becoming ready\n${stderr}`);
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {
      // The child process may still be binding the port.
    }
    await new Promise(resolve => setTimeout(resolve, 50));
  }
  throw new Error(`server did not become ready\n${stderr}`);
}

async function readRequestBody(request) {
  const chunks = [];
  for await (const chunk of request) chunks.push(chunk);
  return Buffer.concat(chunks).toString('utf8');
}

before(async () => {
  const microsoftPort = await availablePort();
  microsoftOrigin = `http://127.0.0.1:${microsoftPort}`;
  microsoftServer = createHttpServer(async (request, response) => {
    if (request.method !== 'POST' || request.url !== '/tenant-id/oauth2/v2.0/token') {
      response.writeHead(404).end();
      return;
    }

    const params = Object.fromEntries(new URLSearchParams(await readRequestBody(request)));
    microsoftTokenRequests.push(params);
    const now = Math.floor(Date.now() / 1000);
    const isRefresh = params.grant_type === 'refresh_token';
    const token = unsignedJwt({
      sub: 'personal-device-test-user',
      preferred_username: 'test@example.cn',
      roles: ['mcp.mail'],
      scp: 'access_as_user',
      exp: now + (isRefresh ? 3600 : 250)
    });
    response.writeHead(200, { 'content-type': 'application/json' });
    response.end(
      JSON.stringify({
        access_token: token,
        refresh_token: 'microsoft-refresh-token',
        id_token: token,
        token_type: 'Bearer',
        scope: 'openid profile offline_access api://api-client-id/access_as_user',
        expires_in: isRefresh ? 3600 : 250
      })
    );
  });
  await new Promise((resolve, reject) => {
    microsoftServer.once('error', reject);
    microsoftServer.listen(microsoftPort, '127.0.0.1', resolve);
  });

  const port = await availablePort();
  origin = `http://127.0.0.1:${port}`;
  statePath = path.join(os.tmpdir(), `21v-mcp-confidential-web-${process.pid}.json`);
  child = spawn(process.execPath, ['dist/index.js'], {
    cwd: process.cwd(),
    env: {
      ...process.env,
      MCP_TRANSPORT: 'http',
      MCP_HTTP_HOST: '127.0.0.1',
      MCP_HTTP_PORT: String(port),
      MCP_HTTP_PATH: '/mcp',
      MCP_HTTP_STATELESS: 'true',
      MCP_PUBLIC_BASE_URL: origin,
      MCP_RESOURCE_URL: '',
      MCP_RESOURCE_METADATA_URL: '',
      MCP_AUTHORIZATION_SERVERS: '',
      MCP_OAUTH_BRIDGE_ISSUER: '',
      MCP_OAUTH_BRIDGE_REDIRECT_URI: '',
      MCP_ALLOW_UNAUTHENTICATED_DISCOVERY: 'false',
      MCP_INBOUND_AUTH_DISABLED: 'false',
      MCP_OAUTH_BRIDGE_ENABLED: 'true',
      MCP_OAUTH_BRIDGE_MICROSOFT_CLIENT_TYPE: 'confidential_web',
      MCP_TOKEN_AUDIENCE: 'api://api-client-id',
      MCP_AUTHORIZATION_SCOPES: 'api://api-client-id/access_as_user',
      MCP_REQUIRED_TOKEN_SCOPES: 'access_as_user',
      MCP_ROLE_BASED_FILTERING: 'false',
      MCP_AUDIT_LOG_ENABLED: 'false',
      MS_TENANT_ID: 'tenant-id',
      MS_CLIENT_ID: 'api-client-id',
      MS_CLIENT_SECRET: 'api-client-secret',
      MS_OAUTH_CLIENT_ID: 'web-client-id',
      MS_OAUTH_CLIENT_SECRET: 'web-client-secret',
      MS_PUBLIC_CLIENT_ID: '',
      MS_AUTHORITY_HOST: microsoftOrigin,
      MCP_OAUTH_BRIDGE_STATE_PATH: statePath
    },
    stdio: ['ignore', 'ignore', 'pipe']
  });
  child.stderr.setEncoding('utf8');
  child.stderr.on('data', chunk => {
    stderr += chunk;
  });
  await waitForServer(`${origin}/healthz`);
});

after(async () => {
  if (child && child.exitCode === null) {
    child.kill('SIGTERM');
    await Promise.race([
      new Promise(resolve => child.once('exit', resolve)),
      new Promise(resolve => setTimeout(resolve, 2_000))
    ]);
    if (child.exitCode === null) child.kill('SIGKILL');
  }
  if (microsoftServer) await new Promise(resolve => microsoftServer.close(resolve));
  if (statePath) {
    await unlink(statePath).catch(error => {
      if (error.code !== 'ENOENT') throw error;
    });
  }
});

test('confidential web mode sends the server secret only to the Microsoft token endpoint', async () => {
  const health = await (await fetch(`${origin}/healthz`)).json();
  assert.equal(health.microsoftLoginClientType, 'confidential_web');

  const clientRedirectUri = 'http://127.0.0.1:54321/callback';
  const registrationResponse = await fetch(`${origin}/oauth/register`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({
      client_name: 'Confidential web integration test',
      redirect_uris: [clientRedirectUri]
    })
  });
  assert.equal(registrationResponse.status, 201);
  const registration = await registrationResponse.json();

  const verifier = 'confidential-web-test-code-verifier-0123456789';
  const challenge = createHash('sha256').update(verifier).digest('base64url');
  const authorizeUrl = new URL(`${origin}/oauth/authorize`);
  authorizeUrl.searchParams.set('response_type', 'code');
  authorizeUrl.searchParams.set('client_id', registration.client_id);
  authorizeUrl.searchParams.set('redirect_uri', clientRedirectUri);
  authorizeUrl.searchParams.set('state', 'client-state');
  authorizeUrl.searchParams.set('code_challenge', challenge);
  authorizeUrl.searchParams.set('code_challenge_method', 'S256');

  const authorizeResponse = await fetch(authorizeUrl, { redirect: 'manual' });
  assert.equal(authorizeResponse.status, 302);
  const microsoftAuthorizeUrl = new URL(authorizeResponse.headers.get('location'));
  assert.equal(microsoftAuthorizeUrl.origin, microsoftOrigin);
  assert.equal(microsoftAuthorizeUrl.searchParams.get('client_id'), 'web-client-id');
  assert.equal(microsoftAuthorizeUrl.searchParams.get('client_secret'), null);
  assert.equal(
    microsoftAuthorizeUrl.searchParams.get('redirect_uri'),
    `${origin}/oauth/microsoft/callback`
  );

  const callbackUrl = new URL(`${origin}/oauth/microsoft/callback`);
  callbackUrl.searchParams.set('code', 'microsoft-authorization-code');
  callbackUrl.searchParams.set('state', microsoftAuthorizeUrl.searchParams.get('state'));
  const callbackResponse = await fetch(callbackUrl, { redirect: 'manual' });
  assert.equal(callbackResponse.status, 302);
  const clientCallback = new URL(callbackResponse.headers.get('location'));
  const bridgeCode = clientCallback.searchParams.get('code');
  assert.match(bridgeCode ?? '', /^mcp_code_/);
  assert.equal(clientCallback.searchParams.get('state'), 'client-state');

  assert.equal(microsoftTokenRequests.length, 1);
  assert.equal(microsoftTokenRequests[0].grant_type, 'authorization_code');
  assert.equal(microsoftTokenRequests[0].client_id, 'web-client-id');
  assert.equal(microsoftTokenRequests[0].client_secret, 'web-client-secret');
  assert.equal(microsoftTokenRequests[0].code, 'microsoft-authorization-code');
  assert.ok(microsoftTokenRequests[0].code_verifier);

  const tokenResponse = await fetch(`${origin}/oauth/token`, {
    method: 'POST',
    headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      grant_type: 'authorization_code',
      client_id: registration.client_id,
      redirect_uri: clientRedirectUri,
      code: bridgeCode,
      code_verifier: verifier
    })
  });
  assert.equal(tokenResponse.status, 200);
  const bridgeToken = await tokenResponse.json();
  assert.match(bridgeToken.access_token, /^mcp_at_/);

  const initializeResponse = await fetch(`${origin}/mcp`, {
    method: 'POST',
    headers: {
      accept: 'application/json, text/event-stream',
      authorization: `Bearer ${bridgeToken.access_token}`,
      'content-type': 'application/json'
    },
    body: JSON.stringify({
      jsonrpc: '2.0',
      id: 1,
      method: 'initialize',
      params: {
        protocolVersion: '2025-06-18',
        capabilities: {},
        clientInfo: { name: 'confidential-web-test', version: '1.0.0' }
      }
    })
  });
  assert.equal(initializeResponse.status, 200);
  assert.equal(microsoftTokenRequests.length, 2);
  assert.equal(microsoftTokenRequests[1].grant_type, 'refresh_token');
  assert.equal(microsoftTokenRequests[1].client_id, 'web-client-id');
  assert.equal(microsoftTokenRequests[1].client_secret, 'web-client-secret');
});
