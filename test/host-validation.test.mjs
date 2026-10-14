import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import http from 'node:http';
import { createRequire } from 'node:module';
import net from 'node:net';
import test from 'node:test';

import { ConfigError, hostHeaderAllowList, loadConfig } from '../dist/config.js';

const { version } = createRequire(import.meta.url)('../package.json');

test('loopback binds accept localhost and the configured public host names', () => {
  const config = {
    httpHost: '127.0.0.1',
    allowedHosts: [],
    resourceUrl: 'https://mcp.example.cn/mcp',
    publicBaseUrl: 'https://mcp.example.cn',
    oauthBridgeIssuer: 'https://login.example.cn'
  };
  const localhostNames = ['localhost', '127.0.0.1', '[::1]'];
  assert.deepEqual(hostHeaderAllowList(config), [...localhostNames, 'mcp.example.cn', 'login.example.cn']);
  assert.deepEqual(hostHeaderAllowList({ ...config, httpHost: 'localhost' }), [
    ...localhostNames,
    'mcp.example.cn',
    'login.example.cn'
  ]);
  assert.equal(hostHeaderAllowList({ ...config, httpHost: '0.0.0.0' }), undefined);
  assert.deepEqual(hostHeaderAllowList({ ...config, httpHost: '0.0.0.0', allowedHosts: ['proxy.example.cn'] }), [
    ...localhostNames,
    'mcp.example.cn',
    'login.example.cn',
    'proxy.example.cn'
  ]);
});

test('MCP_ALLOWED_HOSTS accepts host names, host:port values and URLs', () => {
  const saved = process.env.MCP_ALLOWED_HOSTS;
  try {
    process.env.MCP_ALLOWED_HOSTS = 'Proxy.Example.cn:8443, https://other.example.cn/path [::1] proxy.example.cn';
    assert.deepEqual(loadConfig().allowedHosts, ['proxy.example.cn', 'other.example.cn', '[::1]']);
    process.env.MCP_ALLOWED_HOSTS = 'host^name.example.cn';
    assert.throws(() => loadConfig(), ConfigError);
    delete process.env.MCP_ALLOWED_HOSTS;
    assert.deepEqual(loadConfig().allowedHosts, []);
  } finally {
    if (saved === undefined) delete process.env.MCP_ALLOWED_HOSTS;
    else process.env.MCP_ALLOWED_HOSTS = saved;
  }
});

async function availablePort() {
  const server = net.createServer();
  await new Promise((resolve, reject) => {
    server.once('error', reject);
    server.listen(0, '127.0.0.1', resolve);
  });
  const address = server.address();
  const port = typeof address === 'object' && address ? address.port : 0;
  await new Promise(resolve => server.close(resolve));
  return port;
}

function getHealth(port, host) {
  return new Promise((resolve, reject) => {
    const request = http.request(
      { host: '127.0.0.1', port, path: '/healthz', method: 'GET', headers: host ? { Host: host } : {} },
      response => {
        let body = '';
        response.setEncoding('utf8');
        response.on('data', chunk => {
          body += chunk;
        });
        response.on('end', () => resolve({ status: response.statusCode, body }));
      }
    );
    request.on('error', reject);
    request.end();
  });
}

test('a loopback HTTP server behind a reverse proxy accepts the public Host and rejects others', async (t) => {
  const port = await availablePort();
  let output = '';
  const child = spawn(process.execPath, ['dist/index.js', '--http'], {
    env: {
      ...process.env,
      MCP_HTTP_HOST: '127.0.0.1',
      MCP_HTTP_PORT: String(port),
      MCP_HTTP_PATH: '/mcp',
      MCP_PUBLIC_BASE_URL: 'https://mcp.example.cn',
      MCP_RESOURCE_URL: 'https://mcp.example.cn/mcp',
      MCP_OAUTH_BRIDGE_ISSUER: 'https://mcp.example.cn',
      MCP_OAUTH_HTTP_PATH: '',
      MCP_ALLOWED_HOSTS: '',
      MCP_INBOUND_AUTH_DISABLED: 'true',
      MCP_OAUTH_BRIDGE_ENABLED: 'false',
      MCP_AUDIT_LOG_ENABLED: 'false'
    },
    stdio: ['ignore', 'pipe', 'pipe']
  });
  child.stdout.on('data', chunk => {
    output += chunk.toString();
  });
  child.stderr.on('data', chunk => {
    output += chunk.toString();
  });
  t.after(() => {
    if (child.exitCode === null) child.kill('SIGTERM');
  });

  const deadline = Date.now() + 10_000;
  let local;
  while (Date.now() < deadline) {
    if (child.exitCode !== null) throw new Error(`HTTP test server exited early (${child.exitCode}).\n${output}`);
    local = await getHealth(port).catch(() => undefined);
    if (local?.status === 200) break;
    await new Promise(resolve => setTimeout(resolve, 50));
  }
  assert.equal(local?.status, 200, output);
  assert.equal(JSON.parse(local.body).version, version);

  const proxied = await getHealth(port, 'mcp.example.cn');
  assert.equal(proxied.status, 200);
  const rebinding = await getHealth(port, 'attacker.example.com');
  assert.equal(rebinding.status, 403);
  assert.match(rebinding.body, /Invalid Host/);

  assert.match(output, /Accepted Host headers: localhost, 127\.0\.0\.1, \[::1\], mcp\.example\.cn/);
  assert.doesNotMatch(output, /without DNS rebinding protection/);
});
