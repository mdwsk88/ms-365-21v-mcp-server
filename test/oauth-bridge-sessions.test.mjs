import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';

import express from 'express';

import { loadConfig } from '../dist/config.js';
import { mountOAuthBridgeRoutes } from '../dist/oauth-bridge.js';

const microsoftHost = 'login.partner.microsoftonline.cn';
const day = 24 * 60 * 60;

function unsignedJwt(payload) {
  const encode = value => Buffer.from(JSON.stringify(value)).toString('base64url');
  return `${encode({ alg: 'none', typ: 'JWT' })}.${encode(payload)}.`;
}

/** Answers Microsoft token requests with the next queued reply and passes everything else through. */
function mockMicrosoft(t) {
  const original = globalThis.fetch;
  const replies = [];
  globalThis.fetch = async (url, init) => {
    if (new URL(String(url)).hostname !== microsoftHost) return original(url, init);
    const reply = replies.shift();
    if (!reply) throw new Error('unexpected Microsoft token request');
    if (reply instanceof Error) throw reply;
    return new Response(JSON.stringify(reply.body), { status: reply.status, headers: { 'content-type': 'application/json' } });
  };
  t.after(() => {
    globalThis.fetch = original;
  });
  return replies;
}

const microsoftTokens = (now = Math.floor(Date.now() / 1000)) => ({
  status: 200,
  body: {
    access_token: unsignedJwt({ sub: 'user', oid: 'user-oid', roles: ['mcp.mail'], exp: now + 3600 }),
    refresh_token: 'microsoft-refresh',
    token_type: 'Bearer',
    scope: 'api://api/access_as_user',
    expires_in: 3600
  }
});

function session(name, { expired, refreshToken = 'microsoft-refresh', lastUsedDaysAgo = 0, refreshedDaysAgo }) {
  const now = Math.floor(Date.now() / 1000);
  const saved = {
    clientId: 'client-1',
    accessToken: `${name}-access`,
    refreshToken: `${name}-refresh`,
    scope: ['access_as_user'],
    microsoftToken: {
      accessToken: unsignedJwt({ sub: 'user', oid: 'user-oid', exp: expired ? now - 600 : now + 3600 }),
      refreshToken,
      tokenType: 'Bearer',
      expiresAt: expired ? now - 600 : now + 3600,
      acquiredAt: refreshedDaysAgo === undefined ? now - 7200 : now - refreshedDaysAgo * day
    },
    expiresAt: expired ? now - 600 : now + 3000,
    createdAt: now - 40 * day,
    lastUsedAt: now - lastUsedDaysAgo * day
  };
  // Sessions written by earlier versions have no lastUsedAt.
  if (refreshedDaysAgo !== undefined) delete saved.lastUsedAt;
  return saved;
}

async function bridge(t, sessions) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'mcp-bridge-sessions-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const statePath = path.join(directory, 'state.json');
  const client = {
    clientId: 'client-1',
    redirectUris: ['http://127.0.0.1:5555/callback'],
    grantTypes: ['authorization_code', 'refresh_token'],
    responseTypes: ['code'],
    tokenEndpointAuthMethod: 'none',
    createdAt: Math.floor(Date.now() / 1000) - 50 * day
  };
  await fs.writeFile(statePath, JSON.stringify({ clients: [client], sessions }));
  const config = {
    ...loadConfig(),
    tenantId: 'tenant-id',
    authorityHost: `https://${microsoftHost}`,
    oauthBridgeEnabled: true,
    oauthBridgeMicrosoftClientType: 'confidential_web',
    oauthBridgeMicrosoftClientId: 'web-client-id',
    oauthBridgeMicrosoftClientSecret: 'web-client-secret',
    oauthBridgeIssuer: 'https://mcp.example.cn',
    oauthBridgeRedirectUri: 'https://mcp.example.cn/oauth/microsoft/callback',
    oauthBridgeStatePath: statePath,
    oauthBridgeScopes: ['openid', 'offline_access', 'api://api/access_as_user'],
    oauthBridgeCallbackDelivery: 'redirect',
    publicPathPrefix: '',
    oauthDynamicRegistration: true,
    oauthRedirectHosts: []
  };
  const app = express();
  app.use(express.json());
  app.use(express.urlencoded({ extended: false }));
  mountOAuthBridgeRoutes(app, config);
  const server = app.listen(0, '127.0.0.1');
  await new Promise(resolve => server.once('listening', resolve));
  t.after(() => server.close());
  const origin = `http://127.0.0.1:${server.address().port}`;
  const token = params =>
    fetch(`${origin}/oauth/token`, {
      method: 'POST',
      headers: { 'content-type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams(params)
    });
  const readSessions = async () => JSON.parse(await fs.readFile(statePath, 'utf8')).sessions.map(entry => entry.refreshToken.split('-')[0]);
  return { origin, token, readSessions };
}

test('a new sign-in keeps refreshable sessions and forgets idle or unrefreshable ones', async (t) => {
  const replies = mockMicrosoft(t);
  const { origin, token, readSessions } = await bridge(t, [
    session('overnight', { expired: true, lastUsedDaysAgo: 1 }),
    session('idle', { expired: true, lastUsedDaysAgo: 40 }),
    session('norefresh', { expired: true, refreshToken: null }),
    session('legacy', { expired: true, refreshedDaysAgo: 2 }),
    session('legacystale', { expired: true, refreshedDaysAgo: 45 })
  ]);

  const verifier = 'v'.repeat(43);
  const authorize = new URL(`${origin}/oauth/authorize`);
  for (const [key, value] of Object.entries({
    response_type: 'code',
    client_id: 'client-1',
    redirect_uri: 'http://127.0.0.1:5555/callback',
    code_challenge: createHash('sha256').update(verifier).digest('base64url'),
    code_challenge_method: 'S256'
  })) {
    authorize.searchParams.set(key, value);
  }
  const toMicrosoft = await fetch(authorize, { redirect: 'manual' });
  const state = new URL(toMicrosoft.headers.get('location')).searchParams.get('state');
  replies.push(microsoftTokens());
  const back = await fetch(`${origin}/oauth/microsoft/callback?state=${state}&code=microsoft-code`, { redirect: 'manual' });
  const code = new URL(back.headers.get('location')).searchParams.get('code');
  const granted = await token({
    grant_type: 'authorization_code',
    code,
    client_id: 'client-1',
    redirect_uri: 'http://127.0.0.1:5555/callback',
    code_verifier: verifier
  });
  assert.equal(granted.status, 200);

  const kept = await readSessions();
  assert.equal(kept.length, 3);
  assert.ok(kept.includes('overnight'), 'yesterday’s session can still be refreshed in the morning');
  assert.ok(kept.includes('legacy'), 'a session saved by an earlier version counts its last Microsoft refresh as use');
});

test('refresh asks for a new sign-in only when Microsoft rejects the sign-in for good', async (t) => {
  const replies = mockMicrosoft(t);
  const { token, readSessions } = await bridge(t, [
    session('revoked', { expired: true }),
    session('flaky', { expired: true }),
    session('fine', { expired: true, lastUsedDaysAgo: 3 })
  ]);

  replies.push({ status: 400, body: { error: 'invalid_grant', error_description: 'AADSTS700082: expired' } });
  const revoked = await token({ grant_type: 'refresh_token', refresh_token: 'revoked-refresh', client_id: 'client-1' });
  assert.equal(revoked.status, 400);
  assert.equal((await revoked.json()).error, 'invalid_grant');
  assert.ok(!(await readSessions()).includes('revoked'));

  replies.push(new TypeError('fetch failed'));
  const flaky = await token({ grant_type: 'refresh_token', refresh_token: 'flaky-refresh', client_id: 'client-1' });
  assert.equal(flaky.status, 503);
  assert.equal((await flaky.json()).error, 'temporarily_unavailable');

  replies.push(microsoftTokens());
  const fine = await token({ grant_type: 'refresh_token', refresh_token: 'fine-refresh', client_id: 'client-1' });
  assert.equal(fine.status, 200);
  const body = await fine.json();
  assert.match(body.access_token, /^mcp_at_/);
  assert.equal(body.refresh_token, 'fine-refresh');
  assert.deepEqual((await readSessions()).sort(), ['fine', 'flaky']);

  const wrongClient = await token({ grant_type: 'refresh_token', refresh_token: 'fine-refresh', client_id: 'client-2' });
  assert.equal(wrongClient.status, 400);
});
