#!/usr/bin/env node
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { randomUUID } from 'node:crypto';
import { setTimeout } from 'node:timers/promises';

const image = process.argv[2];
assert.ok(image, 'Usage: node scripts/smoke-container.mjs IMAGE');
const name = `mcp-release-smoke-${randomUUID()}`;
const volume = `${name}-state`;
const publicUrl = 'https://mcp.example.cn';
const appId = '00000000-0000-0000-0000-000000000000';
const docker = (...args) => execFileSync('docker', args, { encoding: 'utf8', timeout: 120_000 }).trim();
const request = (url, options = {}) => fetch(url, { ...options, signal: AbortSignal.timeout(5_000) });

async function ready(base) {
  for (let attempt = 0; attempt < 40; attempt++) {
    try {
      const response = await request(`${base}/healthz`);
      if (response.ok && (await response.json()).ok) return;
    } catch {}
    await setTimeout(500);
  }
  throw new Error('Container did not become healthy');
}

try {
  docker('volume', 'create', volume);
  docker('run', '-d', '--name', name, '-p', '127.0.0.1::3000',
    '--mount', `type=volume,source=${volume},target=/app/.tokens`,
    '-e', `MS_TENANT_ID=${appId}`, '-e', `MS_CLIENT_ID=${appId}`,
    '-e', 'MS_CLIENT_SECRET=replace-container-smoke-placeholder',
    '-e', `MCP_TOKEN_AUDIENCE=api://${appId}`,
    '-e', `MCP_PUBLIC_BASE_URL=${publicUrl}`,
    '-e', 'MCP_OAUTH_BRIDGE_ENABLED=true',
    '-e', 'MCP_OAUTH_BRIDGE_MICROSOFT_CLIENT_TYPE=confidential_web',
    '-e', 'MCP_INBOUND_AUTH_DISABLED=false',
    '-e', 'MCP_ALLOW_UNAUTHENTICATED_DISCOVERY=false', image);
  const port = () => JSON.parse(docker('inspect', name))[0].NetworkSettings.Ports['3000/tcp'][0].HostPort;
  let base = `http://127.0.0.1:${port()}`;
  await ready(base);

  const health = await (await request(`${base}/healthz`)).json();
  assert.equal(health.resourceUrl, `${publicUrl}/mcp`);
  const unauthorized = await request(`${base}/mcp`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Accept: 'application/json, text/event-stream' },
    body: JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'initialize', params: {
      protocolVersion: '2025-03-26', capabilities: {}, clientInfo: { name: 'release-smoke', version: '1' }
    } })
  });
  assert.equal(unauthorized.status, 401, 'Missing bearer must trigger OAuth');
  const metadataUrl = health.resourceMetadataUrl;
  assert.equal(new URL(metadataUrl).origin, publicUrl);
  assert.match(new URL(metadataUrl).pathname, /^\/\.well-known\/oauth-protected-resource(?:\/mcp)?$/);
  assert.ok(unauthorized.headers.get('www-authenticate')?.includes(`resource_metadata="${metadataUrl}"`));
  const resource = await (await request(`${base}${new URL(metadataUrl).pathname}`)).json();
  assert.equal(resource.resource, `${publicUrl}/mcp`);
  assert.deepEqual(resource.authorization_servers, [publicUrl]);
  const authorization = await (await request(`${base}/.well-known/oauth-authorization-server`)).json();
  assert.equal(authorization.issuer, publicUrl);
  for (const key of ['authorization_endpoint', 'token_endpoint', 'registration_endpoint']) {
    assert.equal(new URL(authorization[key]).origin, publicUrl);
  }
  assert.notEqual(docker('exec', name, 'id', '-u'), '0');
  docker('exec', name, 'node', '-e', "require('node:fs').writeFileSync('/app/.tokens/release-smoke', 'persisted')");
  docker('restart', name);
  base = `http://127.0.0.1:${port()}`;
  await ready(base);
  assert.equal(docker('exec', name, 'node', '-e', "process.stdout.write(require('node:fs').readFileSync('/app/.tokens/release-smoke', 'utf8'))"), 'persisted');
  console.log('Container smoke passed: health, OAuth 401/discovery, non-root writable state and restart persistence. No live Entra/Graph requests were made.');
} catch (error) {
  try { console.error(docker('logs', '--tail', '40', name)); } catch {}
  throw error;
} finally {
  try { docker('rm', '-f', name); } catch {}
  try { docker('volume', 'rm', volume); } catch {}
}
