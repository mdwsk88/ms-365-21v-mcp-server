import { randomBytes } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { parseEnv } from 'node:util';

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const ZERO_GUID = '00000000-0000-0000-0000-000000000000';
const LOOPBACK = new Set(['localhost', '127.0.0.1', '[::1]']);
const AUTHORITIES = new Set(['login.chinacloudapi.cn', 'login.partner.microsoftonline.cn']);
const present = value => typeof value === 'string' && value.trim().length > 0;
const placeholder = value => /<[^>]+>|00000000-0000-0000-0000-000000000000|^(?:your[-_ ]|replace[-_ ]|change[-_ ]?me)/i.test(value || '');
const enabled = (value, fallback) => present(value) ? /^(1|true|yes|on)$/i.test(value.trim()) : fallback;
const list = value => (value || '').trim().split(/[\s,]+/).filter(Boolean);
const cleanUrl = url => !url.username && !url.password && !url.search && !url.hash && url.port !== '0' &&
  (url.protocol === 'https:' || (url.protocol === 'http:' && LOOPBACK.has(url.hostname)));

export function normalizePublicBaseUrl(value) {
  if (!present(value) || /[\s#?]/.test(value.trim())) throw new Error('Use an absolute base URL without whitespace, query or fragment.');
  const url = new URL(value.trim());
  if (!cleanUrl(url)) throw new Error('Use HTTPS, or loopback HTTP for local testing; never embed credentials.');
  if (url.pathname.replace(/\/+$/, '').endsWith('/mcp')) throw new Error('Enter the base URL without /mcp.');
  return url.href.replace(/\/+$/, '');
}

/** Single-app starter. It never changes Entra or asks for a secret on the command line. */
export function createQuickstartEnv({ tenantId, clientId, publicBaseUrl = 'http://localhost:3000' }) {
  for (const value of [tenantId, clientId]) {
    if (typeof value !== 'string' || !GUID.test(value) || value === ZERO_GUID) throw new Error('Use non-placeholder tenant and application GUIDs.');
  }
  const base = normalizePublicBaseUrl(publicBaseUrl);
  const entries = {
    MS_TENANT_ID: tenantId,
    MS_CLIENT_ID: clientId,
    MS_CLIENT_SECRET: '<replace-with-client-secret-value>',
    MS_AUTHORITY_HOST: 'https://login.partner.microsoftonline.cn',
    MS_GRAPH_BASE_URL: 'https://microsoftgraph.chinacloudapi.cn/v1.0',
    MCP_TRANSPORT: 'http',
    MCP_HTTP_HOST: '127.0.0.1',
    MCP_HTTP_PORT: '3000',
    MCP_HTTP_PATH: '/mcp',
    MCP_PUBLIC_BASE_URL: base,
    MCP_TOKEN_AUDIENCE: `api://${clientId}`,
    MCP_AUTHORIZATION_SCOPES: `api://${clientId}/access_as_user`,
    MCP_REQUIRED_TOKEN_SCOPES: 'access_as_user',
    MCP_OAUTH_BRIDGE_ENABLED: 'true',
    MCP_OAUTH_BRIDGE_MICROSOFT_CLIENT_TYPE: 'confidential_web',
    MCP_INBOUND_AUTH_DISABLED: 'false',
    MCP_ALLOW_UNAUTHENTICATED_DISCOVERY: 'false',
    MCP_TOOL_CATEGORIES: 'users',
    MCP_DISABLED_GRAPH_SCOPES: 'User.Read.All,User.ReadBasic.All',
    MCP_ROLE_BASED_FILTERING: 'true',
    MCP_TOOL_EXPOSURE_MODE: 'direct',
    MCP_SEND_MODE: 'confirm',
    MCP_CONFIRM_OPERATIONS: 'all',
    MCP_AUDIT_LOG_ENABLED: 'true',
    MCP_ADMIN_TOKEN: randomBytes(32).toString('hex')
  };
  return [
    '# Generated single-app starter. Never commit this file.',
    '# Set MS_CLIENT_SECRET to the secret VALUE, not its ID; quote values containing # or spaces.',
    '# Keep the authority confirmed by your tenant administrator.',
    '# Assign mcp.users to the test user. Start with User.Read only.',
    '# Do not add MCP_RESOURCE_URL: it overrides MCP_PUBLIC_BASE_URL.',
    '# See docs/QUICKSTART.md before enabling other modules or using Docker.',
    ...Object.entries(entries).map(([key, value]) => `${key}=${value}`), ''
  ].join('\n');
}

/** Parse standard .env syntax; inherited environment values take precedence. */
export function readEnvironment(file, inherited = process.env) {
  let fileEnv = {};
  let fileFound = true;
  try { fileEnv = parseEnv(readFileSync(file, 'utf8')); }
  catch (error) {
    if (error.code !== 'ENOENT') throw error;
    fileFound = false;
  }
  return { env: { ...fileEnv, ...inherited }, fileFound };
}

/** Offline HTTP/OAuth preflight. Messages deliberately never interpolate configuration values. */
export function diagnoseEnv(input, { nodeVersion = process.versions.node } = {}) {
  const env = Object.fromEntries(Object.entries(input).filter(([, value]) => typeof value === 'string').map(([key, value]) => [key, value.trim()]));
  const checks = [];
  const add = (level, code, message) => checks.push({ level, code, message });
  const required = key => {
    if (!present(env[key]) || placeholder(env[key])) add('error', key, `Set ${key} to a real value; placeholders are not usable.`);
  };
  if (Number(nodeVersion.split('.')[0]) < 22) add('error', 'NODE_VERSION', 'Use Node.js 22 or newer.');
  for (const key of ['MS_TENANT_ID', 'MS_CLIENT_ID', 'MS_CLIENT_SECRET', 'MCP_TOKEN_AUDIENCE']) required(key);
  for (const key of ['MS_TENANT_ID', 'MS_CLIENT_ID', 'MS_OAUTH_CLIENT_ID']) {
    if (present(env[key]) && !placeholder(env[key]) && !GUID.test(env[key])) add('warning', key, `Check ${key}; the quickstart expects an Entra GUID.`);
  }
  for (const key of ['MCP_OAUTH_BRIDGE_ENABLED', 'MCP_INBOUND_AUTH_DISABLED', 'MCP_ALLOW_UNAUTHENTICATED_DISCOVERY', 'MCP_ROLE_BASED_FILTERING', 'MCP_AUDIT_LOG_ENABLED']) {
    if (present(env[key]) && !/^(true|false|1|0|yes|no|on|off)$/i.test(env[key])) add('error', key, `Use true or false for ${key}.`);
  }
  if (enabled(env.MCP_INBOUND_AUTH_DISABLED, false)) add('error', 'AUTH_DISABLED', 'Re-enable inbound authentication before connecting a client.');
  if (!enabled(env.MCP_ROLE_BASED_FILTERING, true)) add('warning', 'ROLES_DISABLED', 'Enable per-user App Role filtering before sharing the service.');
  if (!enabled(env.MCP_AUDIT_LOG_ENABLED, true)) add('warning', 'AUDIT_DISABLED', 'Audit logging is disabled. Review the deployment policy.');
  if (enabled(env.MCP_ALLOW_UNAUTHENTICATED_DISCOVERY, false)) add('warning', 'ANONYMOUS_DISCOVERY', 'Unauthenticated discovery is enabled. Verify that your platform requires it.');
  if (!enabled(env.MCP_OAUTH_BRIDGE_ENABLED, false)) add('warning', 'BRIDGE_DISABLED', 'The guided client login uses MCP_OAUTH_BRIDGE_ENABLED=true. Other IdP setups need separate verification.');
  else {
    const type = (env.MCP_OAUTH_BRIDGE_MICROSOFT_CLIENT_TYPE || 'public').toLowerCase();
    if (['confidential_web', 'confidential', 'web'].includes(type)) {
      if (present(env.MS_OAUTH_CLIENT_ID) && env.MS_OAUTH_CLIENT_ID !== env.MS_CLIENT_ID) required('MS_OAUTH_CLIENT_SECRET');
      else if (present(env.MS_OAUTH_CLIENT_SECRET)) required('MS_OAUTH_CLIENT_SECRET');
    } else if (type === 'public') required(present(env.MS_OAUTH_CLIENT_ID) ? 'MS_OAUTH_CLIENT_ID' : 'MS_PUBLIC_CLIENT_ID');
    else add('error', 'BRIDGE_CLIENT_TYPE', 'Use public or confidential_web as the bridge client type.');
    if (present(env.MS_OAUTH_CLIENT_ID) && placeholder(env.MS_OAUTH_CLIENT_ID)) add('error', 'MS_OAUTH_CLIENT_ID', 'Remove the placeholder Web client override for a single-app setup.');
  }
  const port = env.MCP_HTTP_PORT || '3000';
  if (!/^\d+$/.test(port) || Number(port) < 1 || Number(port) > 65535) add('error', 'HTTP_PORT', 'MCP_HTTP_PORT must be an integer from 1 to 65535.');
  for (const key of ['MCP_HTTP_PATH', 'MCP_OAUTH_HTTP_PATH']) {
    if (present(env[key]) && (!env[key].startsWith('/') || /[\s?#]/.test(env[key]))) add('error', key, `Use a /path without whitespace, query or fragment for ${key}.`);
  }
  for (const [key, allowed] of [['MCP_TRANSPORT', ['http', 'stdio']], ['MCP_TOOL_EXPOSURE_MODE', ['direct', 'discovery', 'hybrid']], ['MCP_SEND_MODE', ['confirm', 'automatic']]]) {
    if (present(env[key]) && !allowed.includes(env[key].toLowerCase())) add('error', key, `Check the selected mode in ${key}.`);
  }
  if ((env.MCP_TRANSPORT || 'stdio').toLowerCase() !== 'http') add('warning', 'HTTP_PROFILE', 'This preflight targets HTTP/OAuth, not stdio. Use npm run start:http for the guided flow.');
  for (const [key, fallback, hosts] of [
    ['MS_AUTHORITY_HOST', 'https://login.partner.microsoftonline.cn', AUTHORITIES],
    ['MS_GRAPH_BASE_URL', 'https://microsoftgraph.chinacloudapi.cn/v1.0', new Set(['microsoftgraph.chinacloudapi.cn'])]
  ]) {
    try {
      const url = new URL(env[key] || fallback);
      if (!cleanUrl(url) || url.protocol !== 'https:' || !hosts.has(url.hostname) || url.port) throw new Error();
      if (key === 'MS_AUTHORITY_HOST' && url.pathname !== '/') throw new Error();
      if (key === 'MS_GRAPH_BASE_URL' && !['/v1.0', '/v1.0/'].includes(url.pathname)) add('warning', 'GRAPH_VERSION', 'Verify the custom Graph API version; the quickstart targets v1.0.');
    } catch { add('error', key, `Use the documented 21V HTTPS endpoint in ${key}, not a global-cloud endpoint.`); }
  }
  let base;
  if (present(env.MCP_PUBLIC_BASE_URL)) {
    try { base = normalizePublicBaseUrl(env.MCP_PUBLIC_BASE_URL); }
    catch { add('error', 'PUBLIC_BASE_URL', 'Use a clean public base URL without /mcp. Remote hosts require HTTPS.'); }
  } else add('warning', 'PUBLIC_BASE_URL', 'Set MCP_PUBLIC_BASE_URL to derive discovery, issuer, resource and callback URLs consistently.');
  for (const key of ['MCP_PUBLIC_BASE_URL', 'MCP_RESOURCE_URL', 'MCP_RESOURCE_METADATA_URL', 'MCP_OAUTH_BRIDGE_ISSUER', 'MCP_OAUTH_BRIDGE_REDIRECT_URI']) {
    if (!present(env[key])) continue;
    try {
      const url = new URL(env[key]);
      if (placeholder(env[key]) || /(^|\.)example\.(cn|com|org|net)$/.test(url.hostname)) add('error', key, `Replace or remove the example URL in ${key}.`);
      if (!cleanUrl(url)) add('error', key, `Use a clean HTTPS URL in ${key}, or loopback HTTP for local testing.`);
    } catch { add('error', key, `Use an absolute URL in ${key}.`); }
  }
  if (base) {
    const endpoint = (env.MCP_OAUTH_HTTP_PATH || env.MCP_HTTP_PATH || '/mcp').replace(/\/+$/, '');
    for (const [key, expected] of Object.entries({ MCP_RESOURCE_URL: `${base}${endpoint}`, MCP_OAUTH_BRIDGE_ISSUER: base, MCP_OAUTH_BRIDGE_REDIRECT_URI: `${base}/oauth/microsoft/callback` })) {
      if (present(env[key]) && env[key].replace(/\/+$/, '') !== expected) add('warning', `${key}_OVERRIDE`, `${key} overrides MCP_PUBLIC_BASE_URL. Verify proxy routing or remove the override.`);
    }
  }
  if (placeholder(env.MCP_AUTHORIZATION_SCOPES)) add('error', 'AUTHORIZATION_SCOPES', 'Replace the application placeholder in MCP_AUTHORIZATION_SCOPES.');
  if (present(env.MS_CLIENT_ID) && GUID.test(env.MS_CLIENT_ID)) {
    const audiences = list(env.MCP_TOKEN_AUDIENCE);
    if (audiences.length && !audiences.some(value => value === env.MS_CLIENT_ID || value === `api://${env.MS_CLIENT_ID}`)) add('warning', 'TOKEN_AUDIENCE', 'Verify the custom Application ID URI; the audience differs from the usual API application ID.');
    if (present(env.MCP_AUTHORIZATION_SCOPES) && !list(env.MCP_AUTHORIZATION_SCOPES).includes(`api://${env.MS_CLIENT_ID}/access_as_user`)) add('warning', 'API_SCOPE', 'Verify that the scope belongs to the MCP API application, not the Web application.');
  }
  if (!present(env.MCP_ADMIN_TOKEN) || placeholder(env.MCP_ADMIN_TOKEN) || env.MCP_ADMIN_TOKEN.length < 32) add('warning', 'ADMIN_TOKEN', 'Set a long random MCP_ADMIN_TOKEN before using admin endpoints.');
  const errors = checks.filter(check => check.level === 'error').length;
  const warnings = checks.filter(check => check.level === 'warning').length;
  return { ok: errors === 0, errors, warnings, checks, note: 'Offline checks only. Verify Entra consent, App Roles, networking and live OAuth/Graph calls separately.' };
}
