import assert from 'node:assert/strict';
import { test } from 'node:test';
import { spawnSync } from 'node:child_process';
import { cpSync, mkdtempSync, mkdirSync, readFileSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { parseEnv } from 'node:util';
import { createQuickstartEnv, diagnoseEnv, normalizePublicBaseUrl, readEnvironment } from '../scripts/onboarding.mjs';

const tenantId = '11111111-1111-4111-8111-111111111111';
const clientId = '22222222-2222-4222-8222-222222222222';
const otherClientId = '33333333-3333-4333-8333-333333333333';
const options = { tenantId, clientId };
const fixture = () => ({ ...parseEnv(createQuickstartEnv(options)), MS_CLIENT_SECRET: 'fixture-only-not-a-real-secret' });
const codes = report => report.checks.map(check => check.code);
const source = fileURLToPath(new URL('../scripts', import.meta.url));
const sandbox = t => {
  const root = mkdtempSync(path.join(tmpdir(), 'mcp-onboarding-'));
  t.after(() => rmSync(root, { recursive: true, force: true }));
  mkdirSync(path.join(root, 'scripts'));
  for (const name of ['onboarding.mjs', 'setup.mjs', 'doctor.mjs']) cpSync(path.join(source, name), path.join(root, 'scripts', name));
  return root;
};
const run = (root, script, args = [], extraEnv = {}) => {
  const env = Object.fromEntries(Object.entries(process.env).filter(([key]) => !/^(MS_|MCP_)/i.test(key)));
  return spawnSync(process.execPath, [path.join(root, 'scripts', script), ...args], { cwd: root, env: { ...env, ...extraEnv }, encoding: 'utf8', timeout: 10000 });
};

test('starter preserves security controls and requests only profile permissions', () => {
  const env = parseEnv(createQuickstartEnv(options));
  assert.equal(env.MCP_ROLE_BASED_FILTERING, 'true');
  assert.equal(env.MCP_INBOUND_AUTH_DISABLED, 'false');
  assert.equal(env.MCP_ALLOW_UNAUTHENTICATED_DISCOVERY, 'false');
  assert.equal(env.MCP_TOOL_CATEGORIES, 'users');
  assert.equal(env.MCP_DISABLED_GRAPH_SCOPES, 'User.Read.All,User.ReadBasic.All');
  assert.equal(env.MCP_SEND_MODE, 'confirm');
  assert.equal(env.MCP_CONFIRM_OPERATIONS, 'all');
  assert.equal(env.MCP_OAUTH_BRIDGE_MICROSOFT_CLIENT_TYPE, 'confidential_web');
  assert.equal(env.MCP_RESOURCE_URL, undefined);
  assert.equal(env.MS_OAUTH_CLIENT_ID, undefined);
  assert.match(env.MCP_ADMIN_TOKEN, /^[a-f0-9]{64}$/);
  assert.notEqual(env.MCP_ADMIN_TOKEN, parseEnv(createQuickstartEnv(options)).MCP_ADMIN_TOKEN);
});

test('generator rejects placeholder IDs and dotenv injection', () => {
  for (const invalid of ['', '00000000-0000-0000-0000-000000000000', '<TENANT_ID>', tenantId + '\nMCP_INBOUND_AUTH_DISABLED=true']) {
    assert.throws(() => createQuickstartEnv({ ...options, tenantId: invalid }));
  }
});

test('URL normalization preserves deployment prefix and loopback IPv6', () => {
  assert.equal(normalizePublicBaseUrl('https://mcp.test/tools/service/'), 'https://mcp.test/tools/service');
  assert.equal(normalizePublicBaseUrl('http://[::1]:3000/'), 'http://[::1]:3000');
  assert.equal(normalizePublicBaseUrl('http://127.0.0.1:3000'), 'http://127.0.0.1:3000');
});

test('URL normalization rejects remote HTTP, credentials, endpoint suffix and injected values', () => {
  for (const url of ['http://mcp.test', 'https://user:pass@mcp.test', 'https://mcp.test/mcp/', 'https://mcp.test?a=1', 'https://mcp.test/#a', 'https://mcp.test:0', 'https://mcp.test/\nX=1', 'file:///etc/passwd', 'https://mcp.test/?', 'https://mcp.test/#']) {
    assert.throws(() => normalizePublicBaseUrl(url), url);
  }
});

test('starter needs a real secret before preflight passes', () => {
  assert.ok(codes(diagnoseEnv(parseEnv(createQuickstartEnv(options)))).includes('MS_CLIENT_SECRET'));
  assert.deepEqual(diagnoseEnv(fixture()).checks, []);
});

test('doctor detects example URL overrides instead of claiming successful setup', () => {
  const report = diagnoseEnv({ ...fixture(), MCP_RESOURCE_URL: 'https://mcp.example.cn/mcp' });
  assert.equal(report.ok, false);
  assert.ok(codes(report).includes('MCP_RESOURCE_URL'));
  assert.ok(codes(report).includes('MCP_RESOURCE_URL_OVERRIDE'));
});

test('doctor follows single-app fallback and requires separate Web secrets', () => {
  assert.equal(diagnoseEnv({ ...fixture(), MS_OAUTH_CLIENT_ID: clientId }).ok, true);
  assert.ok(codes(diagnoseEnv({ ...fixture(), MS_OAUTH_CLIENT_ID: otherClientId })).includes('MS_OAUTH_CLIENT_SECRET'));
  assert.equal(diagnoseEnv({ ...fixture(), MS_OAUTH_CLIENT_ID: otherClientId, MS_OAUTH_CLIENT_SECRET: 'fixture-web-secret' }).ok, true);
  assert.ok(codes(diagnoseEnv({ ...fixture(), MS_OAUTH_CLIENT_SECRET: '<replace-me>' })).includes('MS_OAUTH_CLIENT_SECRET'));
});

test('doctor understands legacy public client requirements', () => {
  const env = { ...fixture(), MCP_OAUTH_BRIDGE_MICROSOFT_CLIENT_TYPE: 'public' };
  assert.ok(codes(diagnoseEnv(env)).includes('MS_PUBLIC_CLIENT_ID'));
  assert.equal(diagnoseEnv({ ...env, MS_PUBLIC_CLIENT_ID: otherClientId }).ok, true);
});

test('doctor rejects disabled auth and malformed switches and ports', () => {
  assert.ok(codes(diagnoseEnv({ ...fixture(), MCP_INBOUND_AUTH_DISABLED: 'true' })).includes('AUTH_DISABLED'));
  assert.equal(diagnoseEnv({ ...fixture(), MCP_ROLE_BASED_FILTERING: 'flase' }).ok, false);
  for (const port of ['0', '65536', '-1', '3000junk', '3.5']) assert.ok(codes(diagnoseEnv({ ...fixture(), MCP_HTTP_PORT: port })).includes('HTTP_PORT'));
  assert.ok(codes(diagnoseEnv(fixture(), { nodeVersion: '20.0.0' })).includes('NODE_VERSION'));
});

test('doctor warns without changing deliberately selected security policies', () => {
  const env = { ...fixture(), MCP_ROLE_BASED_FILTERING: 'false', MCP_AUDIT_LOG_ENABLED: 'false', MCP_ALLOW_UNAUTHENTICATED_DISCOVERY: 'true' };
  const before = { ...env };
  const report = diagnoseEnv(env);
  assert.equal(report.ok, true);
  assert.equal(report.warnings, 3);
  assert.deepEqual(env, before);
});

test('doctor rejects global endpoints and accepts the two documented tenant authorities', () => {
  assert.equal(diagnoseEnv({ ...fixture(), MS_AUTHORITY_HOST: 'https://login.chinacloudapi.cn' }).ok, true);
  assert.equal(diagnoseEnv({ ...fixture(), MS_AUTHORITY_HOST: 'https://login.microsoftonline.com' }).ok, false);
  assert.equal(diagnoseEnv({ ...fixture(), MS_GRAPH_BASE_URL: 'https://graph.microsoft.com/v1.0' }).ok, false);
});

test('doctor recognizes matching prefixed URLs and custom audiences', () => {
  const env = { ...fixture(), MCP_PUBLIC_BASE_URL: 'https://mcp.test/tools/service', MCP_RESOURCE_URL: 'https://mcp.test/tools/service/mcp', MCP_OAUTH_BRIDGE_ISSUER: 'https://mcp.test/tools/service', MCP_OAUTH_BRIDGE_REDIRECT_URI: 'https://mcp.test/tools/service/oauth/microsoft/callback' };
  assert.deepEqual(diagnoseEnv(env).checks, []);
  assert.ok(codes(diagnoseEnv({ ...env, MCP_TOKEN_AUDIENCE: 'api://custom-api' })).includes('TOKEN_AUDIENCE'));
});

test('diagnostics never print raw credentials, identifiers, malformed URLs or arbitrary input', () => {
  const secret = 'SENSITIVE-fixture-only';
  const report = diagnoseEnv({ ...fixture(), MS_CLIENT_SECRET: secret, MS_OAUTH_CLIENT_SECRET: secret, MCP_ADMIN_TOKEN: secret, MCP_PUBLIC_BASE_URL: `https://${secret}@mcp.test`, MS_AUTHORITY_HOST: secret, MCP_TRANSPORT: secret });
  const output = JSON.stringify(report);
  for (const value of [secret, tenantId, clientId]) assert.equal(output.includes(value), false);
});

test('standard quoted .env values work and environment takes precedence', t => {
  const root = sandbox(t);
  const file = path.join(root, '.env');
  writeFileSync(file, 'MS_CLIENT_SECRET="fixture # with space"\nMCP_HTTP_PORT=3000\n');
  const before = process.env.MCP_HTTP_PORT;
  const result = readEnvironment(file, { MCP_HTTP_PORT: '4000' });
  assert.equal(result.env.MS_CLIENT_SECRET, 'fixture # with space');
  assert.equal(result.env.MCP_HTTP_PORT, '4000');
  assert.equal(process.env.MCP_HTTP_PORT, before);
  assert.equal(readEnvironment(path.join(root, 'missing'), {}).fileFound, false);
});

test('setup CLI generates a protected file but never overwrites it', t => {
  const root = sandbox(t);
  const args = ['--tenant-id', tenantId, '--client-id', clientId];
  const first = run(root, 'setup.mjs', args);
  assert.equal(first.status, 0, first.stderr);
  const file = path.join(root, '.env');
  const contents = readFileSync(file, 'utf8');
  const env = parseEnv(contents);
  assert.equal(first.stdout.includes(env.MCP_ADMIN_TOKEN), false);
  if (process.platform !== 'win32') assert.equal(statSync(file).mode & 0o777, 0o600);
  const second = run(root, 'setup.mjs', args);
  assert.equal(second.status, 1);
  assert.match(second.stderr, /already exists/);
  assert.equal(readFileSync(file, 'utf8'), contents);
});

test('setup fails fast without a TTY and does not echo invalid arguments', t => {
  const root = sandbox(t);
  const missing = run(root, 'setup.mjs');
  assert.equal(missing.status, 1);
  const invalid = run(root, 'setup.mjs', ['--client-secret', 'SENSITIVE-fixture-only']);
  assert.equal(invalid.status, 1);
  assert.equal((invalid.stdout + invalid.stderr).includes('SENSITIVE-fixture-only'), false);
});

test('doctor CLI provides clean JSON, exit codes and inherited environment support', t => {
  const root = sandbox(t);
  const failed = run(root, 'doctor.mjs', ['--json']);
  assert.equal(failed.status, 1);
  assert.equal(JSON.parse(failed.stdout).ok, false);
  const passed = run(root, 'doctor.mjs', ['--json'], fixture());
  assert.equal(passed.status, 0, passed.stderr);
  assert.equal(JSON.parse(passed.stdout).ok, true);
  assert.ok(codes(JSON.parse(passed.stdout)).includes('ENV_FILE_MISSING'));
  assert.equal(passed.stdout.includes('fixture-only-not-a-real-secret'), false);
  const unreadable = run(root, 'doctor.mjs', ['--json', '--config', root]);
  assert.equal(unreadable.status, 1);
  assert.ok(codes(JSON.parse(unreadable.stdout)).includes('DOCTOR_INPUT'));
});
