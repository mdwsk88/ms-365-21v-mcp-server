import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';

import { loadConfig } from '../dist/config.js';
import { authenticateOAuthBridgeToken } from '../dist/oauth-bridge.js';

function unsignedJwt(payload) {
  const encode = (value) => Buffer.from(JSON.stringify(value)).toString('base64url');
  return `${encode({ alg: 'none', typ: 'JWT' })}.${encode(payload)}.`;
}

test('OAuth bridge retains Microsoft Entra app roles in its persisted session and request claims', async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'mcp-bridge-role-'));
  const statePath = path.join(directory, 'state.json');
  const now = Math.floor(Date.now() / 1000);
  const microsoftAccessToken = unsignedJwt({
    sub: 'microsoft-user',
    roles: ['mcp.mail', 'mcp.calendar'],
    exp: now + 3600
  });
  const microsoftIdToken = unsignedJwt({
    sub: 'microsoft-user',
    preferred_username: 'user@example.cn',
    roles: ['mcp.admin'],
    exp: now + 3600
  });
  await fs.writeFile(
    statePath,
    JSON.stringify({
      clients: [],
      sessions: [
        {
          clientId: 'dynamic-client',
          accessToken: 'bridge-access-token',
          refreshToken: 'bridge-refresh-token',
          scope: ['access_as_user'],
          microsoftToken: {
            accessToken: microsoftAccessToken,
            idToken: microsoftIdToken,
            tokenType: 'Bearer',
            scopes: 'access_as_user',
            expiresAt: now + 3600,
            acquiredAt: now
          },
          expiresAt: now + 3500,
          createdAt: now
        }
      ]
    })
  );

  const config = {
    ...loadConfig(),
    oauthBridgeEnabled: true,
    oauthBridgeStatePath: statePath,
    requiredTokenScopes: ['access_as_user'],
    authorizationScopes: ['api://gateway/access_as_user']
  };

  const result = await authenticateOAuthBridgeToken(config, 'bridge-access-token');
  assert.deepEqual(result?.claims.roles, ['mcp.mail', 'mcp.calendar']);
  assert.equal(result?.claims.sub, 'microsoft-user');
  assert.equal(result?.claims.preferred_username, 'user@example.cn');

  const persisted = JSON.parse(await fs.readFile(statePath, 'utf8'));
  assert.deepEqual(persisted.sessions[0].roles, ['mcp.mail', 'mcp.calendar']);
});

test('OAuth bridge serializes concurrent state persistence without temp-file collisions', async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'mcp-bridge-concurrent-'));
  const statePath = path.join(directory, 'state.json');
  const now = Math.floor(Date.now() / 1000);
  const microsoftAccessToken = unsignedJwt({
    sub: 'concurrent-user',
    roles: ['mcp.mail'],
    exp: now + 3600
  });
  await fs.writeFile(
    statePath,
    JSON.stringify({
      clients: [],
      sessions: [
        {
          clientId: 'concurrent-client',
          accessToken: 'concurrent-bridge-access-token',
          refreshToken: 'concurrent-bridge-refresh-token',
          scope: ['access_as_user'],
          microsoftToken: {
            accessToken: microsoftAccessToken,
            tokenType: 'Bearer',
            scopes: 'access_as_user',
            expiresAt: now + 3600,
            acquiredAt: now
          },
          expiresAt: now + 3500,
          createdAt: now
        }
      ]
    })
  );

  const config = {
    ...loadConfig(),
    oauthBridgeEnabled: true,
    oauthBridgeStatePath: statePath,
    requiredTokenScopes: ['access_as_user'],
    authorizationScopes: ['api://gateway/access_as_user']
  };

  const results = await Promise.all(
    Array.from({ length: 32 }, () =>
      authenticateOAuthBridgeToken(config, 'concurrent-bridge-access-token')
    )
  );
  assert.equal(results.filter(Boolean).length, 32);

  const persisted = JSON.parse(await fs.readFile(statePath, 'utf8'));
  assert.equal(persisted.sessions.length, 1);
  assert.deepEqual(persisted.sessions[0].roles, ['mcp.mail']);
  assert.deepEqual(
    (await fs.readdir(directory)).filter(name => name.endsWith('.tmp')),
    []
  );
});
