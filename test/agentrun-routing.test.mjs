import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { unlink } from 'node:fs/promises';
import { createServer } from 'node:net';
import { after, before, test } from 'node:test';
import { discoverOAuthServerInfo } from '@modelcontextprotocol/sdk/client/auth.js';

const publicBaseUrl = 'https://gateway.example/tools/MCP_21V';
let child;
let origin;
let stderr = '';
const statePath = `/tmp/21v-mcp-test-${process.pid}.json`;

async function availablePort() {
  const server = createServer();
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
    if (child.exitCode !== null) {
      throw new Error(`server exited before becoming ready\n${stderr}`);
    }
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {
      // The process may still be binding the port.
    }
    await new Promise(resolve => setTimeout(resolve, 50));
  }
  throw new Error(`server did not become ready\n${stderr}`);
}

before(async () => {
  const port = await availablePort();
  origin = `http://127.0.0.1:${port}`;
  child = spawn(process.execPath, ['dist/index.js'], {
    cwd: process.cwd(),
    env: {
      ...process.env,
      MCP_TRANSPORT: 'http',
      MCP_HTTP_HOST: '127.0.0.1',
      MCP_HTTP_PORT: String(port),
      MCP_HTTP_PATH: '/mcp',
      MCP_OAUTH_HTTP_PATH: '/oauth-mcp',
      MCP_HTTP_STATELESS: 'true',
      MCP_PUBLIC_BASE_URL: publicBaseUrl,
      MCP_RESOURCE_URL: '',
      MCP_RESOURCE_METADATA_URL: '',
      MCP_AUTHORIZATION_SERVERS: '',
      MCP_OAUTH_BRIDGE_ISSUER: '',
      MCP_OAUTH_BRIDGE_REDIRECT_URI: '',
      MCP_ALLOW_UNAUTHENTICATED_DISCOVERY: 'true',
      MCP_INBOUND_AUTH_DISABLED: 'false',
      MCP_OAUTH_BRIDGE_ENABLED: 'true',
      MCP_TOKEN_AUDIENCE: 'api://test-mcp-app',
      MCP_AUTHORIZATION_SCOPES: 'api://test-mcp-app/access_as_user',
      MCP_REQUIRED_TOKEN_SCOPES: 'access_as_user',
      MS_TENANT_ID: '00000000-0000-0000-0000-000000000001',
      MS_CLIENT_ID: '00000000-0000-0000-0000-000000000002',
      MS_CLIENT_SECRET: 'test-only-secret',
      MS_PUBLIC_CLIENT_ID: '00000000-0000-0000-0000-000000000003',
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
  await unlink(statePath).catch(error => {
    if (error.code !== 'ENOENT') throw error;
  });
});

test('derives a complete public OAuth surface from one AgentRun base URL', async () => {
  const resourceResponse = await fetch(`${origin}/tools/MCP_21V/.well-known/oauth-protected-resource`);
  assert.equal(resourceResponse.status, 200);
  assert.deepEqual(await resourceResponse.json(), {
    resource: `${publicBaseUrl}/oauth-mcp`,
    authorization_servers: [publicBaseUrl],
    bearer_methods_supported: ['header'],
    scopes_supported: ['api://test-mcp-app/access_as_user'],
    resource_name: '21V Microsoft Graph MCP'
  });

  const authorizationResponse = await fetch(
    `${origin}/tools/MCP_21V/.well-known/oauth-authorization-server`
  );
  assert.equal(authorizationResponse.status, 200);
  const authorization = await authorizationResponse.json();
  assert.equal(authorization.issuer, publicBaseUrl);
  assert.equal(authorization.authorization_endpoint, `${publicBaseUrl}/oauth/authorize`);
  assert.equal(authorization.token_endpoint, `${publicBaseUrl}/oauth/token`);
  assert.equal(authorization.registration_endpoint, `${publicBaseUrl}/oauth/register`);

  const openIdResponse = await fetch(
    `${origin}/tools/MCP_21V/.well-known/openid-configuration`
  );
  assert.equal(openIdResponse.status, 200);
  const openId = await openIdResponse.json();
  assert.equal(openId.jwks_uri, `${publicBaseUrl}/.well-known/jwks.json`);
  assert.deepEqual(openId.subject_types_supported, ['public']);
  assert.deepEqual(openId.id_token_signing_alg_values_supported, ['RS256']);

  const jwksResponse = await fetch(`${origin}/tools/MCP_21V/.well-known/jwks.json`);
  assert.equal(jwksResponse.status, 200);
  assert.deepEqual(await jwksResponse.json(), { keys: [] });

  const callbackResponse = await fetch(`${origin}/tools/MCP_21V/oauth/microsoft/callback`, {
    redirect: 'manual'
  });
  assert.equal(callbackResponse.status, 400);
  assert.equal((await callbackResponse.json()).error, 'invalid_request');
});

test('supports Qoder path-aware OAuth discovery through the AgentRun prefix', async () => {
  const gatewayOrigin = new URL(publicBaseUrl).origin;
  const fetchThroughAgentRun = (input, init) => {
    const requested = new URL(input instanceof Request ? input.url : input.toString());
    if (requested.origin === gatewayOrigin) {
      requested.protocol = 'http:';
      requested.host = new URL(origin).host;
    }
    return fetch(requested, init);
  };

  const serverInfo = await discoverOAuthServerInfo(`${publicBaseUrl}/oauth-mcp`, {
    resourceMetadataUrl: new URL(`${publicBaseUrl}/.well-known/oauth-protected-resource`),
    fetchFn: fetchThroughAgentRun
  });

  assert.equal(serverInfo.authorizationServerUrl, publicBaseUrl);
  assert.equal(
    serverInfo.authorizationServerMetadata.authorization_endpoint,
    `${publicBaseUrl}/oauth/authorize`
  );
  assert.equal(
    serverInfo.authorizationServerMetadata.token_endpoint,
    `${publicBaseUrl}/oauth/token`
  );
  assert.equal(
    serverInfo.authorizationServerMetadata.registration_endpoint,
    `${publicBaseUrl}/oauth/register`
  );
});

test('keeps the client callback dynamic while Microsoft uses the fixed AgentRun callback', async () => {
  const clientRedirectUri = 'http://127.0.0.1:54321/callback';
  const registrationResponse = await fetch(`${origin}/tools/MCP_21V/oauth/register`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({
      client_name: 'AgentRun OAuth routing test',
      redirect_uris: [clientRedirectUri]
    })
  });
  assert.equal(registrationResponse.status, 201);
  const registration = await registrationResponse.json();
  assert.match(registration.client_id, /^mcp_client_/);

  const authorizeUrl = new URL(`${origin}/tools/MCP_21V/oauth/authorize`);
  authorizeUrl.searchParams.set('response_type', 'code');
  authorizeUrl.searchParams.set('client_id', registration.client_id);
  authorizeUrl.searchParams.set('redirect_uri', clientRedirectUri);
  authorizeUrl.searchParams.set('state', 'client-state');
  authorizeUrl.searchParams.set('code_challenge', 'test-code-challenge');
  authorizeUrl.searchParams.set('code_challenge_method', 'S256');

  const authorizeResponse = await fetch(authorizeUrl, { redirect: 'manual' });
  assert.equal(authorizeResponse.status, 302);
  const microsoftUrl = new URL(authorizeResponse.headers.get('location'));
  assert.equal(microsoftUrl.origin, 'https://login.partner.microsoftonline.cn');
  assert.equal(
    microsoftUrl.pathname,
    '/00000000-0000-0000-0000-000000000001/oauth2/v2.0/authorize'
  );
  assert.equal(microsoftUrl.searchParams.get('client_id'), '00000000-0000-0000-0000-000000000003');
  assert.equal(
    microsoftUrl.searchParams.get('redirect_uri'),
    `${publicBaseUrl}/oauth/microsoft/callback`
  );
  assert.match(microsoftUrl.searchParams.get('state') ?? '', /^ms_state_/);
});

test('separates AgentRun catalog discovery from the OAuth-required client endpoint', async () => {
  const initializeResponse = await fetch(`${origin}/tools/MCP_21V/mcp`, {
    method: 'POST',
    headers: {
      accept: 'application/json, text/event-stream',
      'content-type': 'application/json'
    },
    body: JSON.stringify({
      jsonrpc: '2.0',
      id: 1,
      method: 'initialize',
      params: {
        protocolVersion: '2025-06-18',
        capabilities: {},
        clientInfo: { name: 'agentrun-routing-test', version: '1.0.0' }
      }
    })
  });

  assert.equal(initializeResponse.status, 200);
  assert.equal(initializeResponse.headers.get('mcp-session-id'), null);

  const protectedInitializeResponse = await fetch(`${origin}/tools/MCP_21V/oauth-mcp`, {
    method: 'POST',
    headers: {
      accept: 'application/json, text/event-stream',
      'content-type': 'application/json'
    },
    body: JSON.stringify({
      jsonrpc: '2.0',
      id: 2,
      method: 'initialize',
      params: {
        protocolVersion: '2025-06-18',
        capabilities: {},
        clientInfo: { name: 'qoder-oauth-routing-test', version: '1.0.0' }
      }
    })
  });

  assert.equal(protectedInitializeResponse.status, 401);
  assert.match(
    protectedInitializeResponse.headers.get('www-authenticate') ?? '',
    new RegExp(`resource_metadata="${publicBaseUrl}/\\.well-known/oauth-protected-resource"`)
  );

  const toolCallResponse = await fetch(`${origin}/tools/MCP_21V/mcp`, {
    method: 'POST',
    headers: {
      accept: 'application/json, text/event-stream',
      'content-type': 'application/json'
    },
    body: JSON.stringify({
      jsonrpc: '2.0',
      id: 3,
      method: 'tools/call',
      params: { name: 'graph_get_me', arguments: {} }
    })
  });

  assert.equal(toolCallResponse.status, 401);
  assert.match(
    toolCallResponse.headers.get('www-authenticate') ?? '',
    new RegExp(`resource_metadata="${publicBaseUrl}/\\.well-known/oauth-protected-resource"`)
  );
});
