import path from 'node:path';
import { fileURLToPath } from 'node:url';
import dotenv from 'dotenv';
import { parseDownstreamServices, type DownstreamMcp } from './gateway/proxy.js';

export type GraphResilienceConfig = {
  maxRetries: number;
  initialBackoffMs: number;
  backoffMultiplier: number;
  circuitBreakerThreshold: number;
  circuitBreakerCooldownMs: number;
  timeoutMs: number;
};

export type OAuthBridgeMicrosoftClientType = 'public' | 'confidential_web';
export type SendMode = 'confirm' | 'automatic';
export type ToolExposureMode = 'direct' | 'discovery' | 'hybrid';

export type AppConfig = {
  tenantId?: string;
  clientId?: string;
  clientSecret?: string;
  publicClientId?: string;
  oauthBridgeMicrosoftClientType: OAuthBridgeMicrosoftClientType;
  oauthBridgeMicrosoftClientId?: string;
  oauthBridgeMicrosoftClientSecret?: string;
  authorityHost: string;
  graphBaseUrl: string;
  graphResource: string;
  scopes: string[];
  oboScopes: string[];
  tokenCachePath: string;
  deviceCodeCachePath: string;
  enableRawGraphGet: boolean;
  transport: 'stdio' | 'http';
  httpPort: number;
  httpHost: string;
  httpPath: string;
  oauthHttpPath?: string;
  httpStateless: boolean;
  publicBaseUrl?: string;
  publicPathPrefix: string;
  resourceUrl: string;
  resourceMetadataUrl: string;
  authorizationServers: string[];
  authorizationScopes: string[];
  requiredTokenScopes: string[];
  resourceDocumentationUrl?: string;
  inboundAuthDisabled: boolean;
  allowUnauthenticatedDiscovery: boolean;
  inboundTokenAudiences: string[];
  inboundTokenIssuer?: string;
  oidcMetadataUrl: string;
  oauthBridgeEnabled: boolean;
  oauthBridgeIssuer: string;
  oauthBridgeRedirectUri: string;
  oauthBridgeScopes: string[];
  oauthBridgePrompt?: string;
  oauthBridgeCallbackDelivery: 'redirect' | 'background';
  oauthBridgeStatePath: string;
  toolExposureMode: ToolExposureMode;
  toolCategories: string[];
  directToolCategories: string[];
  directTools: string[];
  disabledGraphScopes: string[];
  roleBasedFiltering: boolean;
  auditLogEnabled: boolean;
  auditLogPath: string;
  sendMode: SendMode;
  confirmOperations: string[];
  confirmTtlSeconds: number;
  graphResilience: GraphResilienceConfig;
  adminToken?: string;
  downstreamServices: DownstreamMcp[];
  missing: string[];
  missingRemote: string[];
};

export class ConfigError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'ConfigError';
  }
}

const packageRoot = fileURLToPath(new URL('..', import.meta.url));
dotenv.config({ path: path.join(packageRoot, '.env'), override: false });

function stripTrailingSlash(value: string): string {
  return value.replace(/\/+$/, '');
}

function stripTrailingSlashExceptRoot(value: string): string {
  const stripped = value.replace(/\/+$/, '');
  return stripped || '/';
}

function optionalEnv(name: string): string | undefined {
  const value = process.env[name]?.trim();
  return value ? value : undefined;
}

function booleanEnv(name: string, fallback: boolean): boolean {
  const value = optionalEnv(name);
  if (!value) return fallback;
  return ['1', 'true', 'yes', 'on'].includes(value.toLowerCase());
}

function integerEnv(name: string, fallback: number): number {
  const value = optionalEnv(name);
  if (!value) return fallback;
  const parsed = Number.parseInt(value, 10);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function positiveIntegerEnv(name: string, fallback: number): number {
  const parsed = integerEnv(name, fallback);
  return parsed >= 0 ? parsed : fallback;
}

function resolvePath(value: string | undefined, fallbackFromPackageRoot: string): string {
  if (!value) return path.join(packageRoot, fallbackFromPackageRoot);
  return path.isAbsolute(value) ? value : path.resolve(process.cwd(), value);
}

function graphResourceFromBaseUrl(graphBaseUrl: string): string {
  const url = new URL(graphBaseUrl);
  return `${url.protocol}//${url.host}`;
}

function parseList(value: string | undefined): string[] {
  if (!value) return [];
  return value
    .split(/[\s,]+/)
    .map(item => item.trim())
    .filter(Boolean);
}

function normalizeHttpPath(value: string | undefined): string {
  const pathValue = value?.trim() || '/mcp';
  return stripTrailingSlashExceptRoot(pathValue.startsWith('/') ? pathValue : `/${pathValue}`);
}

function publicPathPrefixFromUrl(value: string): string {
  const pathname = stripTrailingSlashExceptRoot(new URL(value).pathname);
  return pathname === '/' ? '' : pathname;
}

function appendUrlPath(baseUrl: string, pathname: string): string {
  const normalizedPath = pathname.startsWith('/') ? pathname : `/${pathname}`;
  return `${stripTrailingSlash(baseUrl)}${normalizedPath}`;
}

function localResourceUrl(httpHost: string, httpPort: number, httpPath: string): string {
  const host = httpHost === '0.0.0.0' || httpHost === '::' ? '127.0.0.1' : httpHost;
  return `http://${host}:${httpPort}${httpPath}`;
}

function resourceMetadataUrlFromResource(resourceUrl: string): string {
  const url = new URL(resourceUrl);
  const path = stripTrailingSlashExceptRoot(url.pathname);
  const suffix = path === '/' ? '' : path;
  return `${url.origin}/.well-known/oauth-protected-resource${suffix}`;
}

function authorizationServerFromTenant(authorityHost: string, tenantId?: string): string {
  return tenantId ? `${authorityHost}/${encodeURIComponent(tenantId)}/v2.0` : authorityHost;
}

function localAuthorizationServer(resourceUrl: string): string {
  const url = new URL(resourceUrl);
  return url.origin;
}

function scopeBaseFromAudience(audience: string): string {
  if (/^https?:\/\//i.test(audience) || audience.startsWith('api://')) {
    return stripTrailingSlash(audience);
  }
  return `api://${audience}`;
}

function defaultAuthorizationScopes(audiences: string[]): string[] {
  const audience = audiences[0];
  if (!audience) return [];
  const base = scopeBaseFromAudience(audience);
  return [`${base}/access_as_user`];
}

function scopeName(scope: string): string {
  const trimmed = stripTrailingSlash(scope.trim());
  const slashIndex = trimmed.lastIndexOf('/');
  return slashIndex >= 0 ? trimmed.slice(slashIndex + 1) : trimmed;
}

function defaultRequiredTokenScopes(authorizationScopes: string[]): string[] {
  return authorizationScopes.map(scopeName).filter(Boolean);
}

function parseScopes(value: string | undefined, graphResource: string): string[] {
  if (!value) {
    return ['openid', 'profile', 'offline_access', `${graphResource}/User.Read`];
  }
  return parseList(value);
}

function parseOboScopes(value: string | undefined, graphResource: string): string[] {
  const scopes = parseList(value);
  return scopes.length ? scopes : [`${graphResource}/User.Read`];
}

function parseTransport(): 'stdio' | 'http' {
  if (process.argv.includes('--http')) return 'http';
  if (process.argv.includes('--stdio')) return 'stdio';

  const value = optionalEnv('MCP_TRANSPORT')?.toLowerCase();
  return value === 'http' ? 'http' : 'stdio';
}

function parseOAuthBridgeCallbackDelivery(): 'redirect' | 'background' {
  const value = optionalEnv('MCP_OAUTH_BRIDGE_CALLBACK_DELIVERY')?.toLowerCase();
  return value === 'background' ? 'background' : 'redirect';
}

function parseOAuthBridgeMicrosoftClientType(): OAuthBridgeMicrosoftClientType {
  const value = optionalEnv('MCP_OAUTH_BRIDGE_MICROSOFT_CLIENT_TYPE')?.toLowerCase();
  if (!value || value === 'public') return 'public';
  if (value === 'confidential_web' || value === 'confidential' || value === 'web') {
    return 'confidential_web';
  }
  throw new ConfigError(
    'MCP_OAUTH_BRIDGE_MICROSOFT_CLIENT_TYPE must be public or confidential_web.'
  );
}

function parseSendMode(): SendMode {
  const value = optionalEnv('MCP_SEND_MODE')?.toLowerCase();
  if (!value || value === 'confirm') return 'confirm';
  if (value === 'automatic') return 'automatic';
  throw new ConfigError('MCP_SEND_MODE must be confirm or automatic.');
}

function parseToolExposureMode(): ToolExposureMode {
  const value = optionalEnv('MCP_TOOL_EXPOSURE_MODE')?.toLowerCase();
  if (!value || value === 'direct') return 'direct';
  if (value === 'discovery' || value === 'hybrid') return value;
  throw new ConfigError('MCP_TOOL_EXPOSURE_MODE must be direct, discovery, or hybrid.');
}

export function loadConfig(): AppConfig {
  const tenantId = optionalEnv('MS_TENANT_ID');
  const clientId = optionalEnv('MS_CLIENT_ID');
  const clientSecret = optionalEnv('MS_CLIENT_SECRET');
  const publicClientId = optionalEnv('MS_PUBLIC_CLIENT_ID');
  const oauthClientIdOverride = optionalEnv('MS_OAUTH_CLIENT_ID');
  const oauthClientSecretOverride = optionalEnv('MS_OAUTH_CLIENT_SECRET');
  const oauthBridgeMicrosoftClientType = parseOAuthBridgeMicrosoftClientType();
  const oauthBridgeMicrosoftClientId =
    oauthClientIdOverride ??
    (oauthBridgeMicrosoftClientType === 'confidential_web' ? clientId : publicClientId);
  const oauthBridgeMicrosoftClientSecret =
    oauthBridgeMicrosoftClientType === 'confidential_web'
      ? oauthClientSecretOverride ??
        (!oauthClientIdOverride || oauthClientIdOverride === clientId ? clientSecret : undefined)
      : undefined;
  const authorityHost = stripTrailingSlash(
    optionalEnv('MS_AUTHORITY_HOST') ?? 'https://login.partner.microsoftonline.cn'
  );
  const graphBaseUrl = stripTrailingSlash(
    optionalEnv('MS_GRAPH_BASE_URL') ?? 'https://microsoftgraph.chinacloudapi.cn/v1.0'
  );
  const graphResource = graphResourceFromBaseUrl(graphBaseUrl);
  const tokenCachePath = resolvePath(optionalEnv('MS_TOKEN_CACHE_PATH'), '.tokens/21v-graph-token.json');
  const deviceCodeCachePath = `${tokenCachePath}.device-code`;
  const transport = parseTransport();
  const httpPort = integerEnv('MCP_HTTP_PORT', 3000);
  const httpHost = optionalEnv('MCP_HTTP_HOST') ?? '127.0.0.1';
  const httpPath = normalizeHttpPath(optionalEnv('MCP_HTTP_PATH'));
  const oauthHttpPathValue = optionalEnv('MCP_OAUTH_HTTP_PATH');
  const oauthHttpPath = oauthHttpPathValue ? normalizeHttpPath(oauthHttpPathValue) : undefined;
  const protectedHttpPath = oauthHttpPath ?? httpPath;
  const httpStateless = booleanEnv('MCP_HTTP_STATELESS', false);
  const publicBaseUrlValue = optionalEnv('MCP_PUBLIC_BASE_URL');
  const publicBaseUrl = publicBaseUrlValue ? stripTrailingSlash(publicBaseUrlValue) : undefined;
  const inboundTokenAudiences = parseList(optionalEnv('MCP_TOKEN_AUDIENCE'));
  const resourceUrl = stripTrailingSlashExceptRoot(
    optionalEnv('MCP_RESOURCE_URL') ??
      (publicBaseUrl
        ? appendUrlPath(publicBaseUrl, protectedHttpPath)
        : localResourceUrl(httpHost, httpPort, protectedHttpPath))
  );
  const resourceMetadataUrl =
    optionalEnv('MCP_RESOURCE_METADATA_URL') ??
    (publicBaseUrl
      ? appendUrlPath(publicBaseUrl, '/.well-known/oauth-protected-resource')
      : resourceMetadataUrlFromResource(resourceUrl));
  const oauthBridgeEnabled = booleanEnv('MCP_OAUTH_BRIDGE_ENABLED', false);
  const authorizationServers = parseList(optionalEnv('MCP_AUTHORIZATION_SERVERS'));
  if (authorizationServers.length === 0) {
    authorizationServers.push(
      oauthBridgeEnabled
        ? (publicBaseUrl ?? localAuthorizationServer(resourceUrl))
        : authorizationServerFromTenant(authorityHost, tenantId)
    );
  }
  const authorizationScopes = parseList(optionalEnv('MCP_AUTHORIZATION_SCOPES'));
  if (authorizationScopes.length === 0) {
    authorizationScopes.push(...defaultAuthorizationScopes(inboundTokenAudiences));
  }
  const requiredTokenScopes = parseList(optionalEnv('MCP_REQUIRED_TOKEN_SCOPES'));
  if (requiredTokenScopes.length === 0) {
    requiredTokenScopes.push(...defaultRequiredTokenScopes(authorizationScopes));
  }
  const inboundAuthDisabled = booleanEnv('MCP_INBOUND_AUTH_DISABLED', false);
  const allowUnauthenticatedDiscovery = booleanEnv('MCP_ALLOW_UNAUTHENTICATED_DISCOVERY', false);
  const oidcMetadataUrl =
    optionalEnv('MS_OIDC_METADATA_URL') ??
    (tenantId ? `${authorityHost}/${encodeURIComponent(tenantId)}/v2.0/.well-known/openid-configuration` : '');
  const oauthBridgeIssuer = stripTrailingSlash(
    optionalEnv('MCP_OAUTH_BRIDGE_ISSUER') ?? publicBaseUrl ?? localAuthorizationServer(resourceUrl)
  );
  const oauthBridgeRedirectUri =
    optionalEnv('MCP_OAUTH_BRIDGE_REDIRECT_URI') ?? `${oauthBridgeIssuer}/oauth/microsoft/callback`;
  const oauthBridgeScopes = parseList(optionalEnv('MCP_OAUTH_BRIDGE_SCOPES'));
  if (oauthBridgeScopes.length === 0) {
    oauthBridgeScopes.push('openid', 'profile', 'offline_access', ...authorizationScopes);
  }
  const oauthBridgePrompt = optionalEnv('MCP_OAUTH_BRIDGE_PROMPT');
  const oauthBridgeCallbackDelivery = parseOAuthBridgeCallbackDelivery();
  const oauthBridgeStatePath = resolvePath(
    optionalEnv('MCP_OAUTH_BRIDGE_STATE_PATH'),
    '.tokens/oauth-bridge-state.json'
  );
  const publicPathPrefix = publicPathPrefixFromUrl(publicBaseUrl ?? oauthBridgeIssuer);
  const toolExposureMode = parseToolExposureMode();
  const toolCategories = [...new Set(parseList(optionalEnv('MCP_TOOL_CATEGORIES')).map(value => value.toLowerCase()))];
  const directToolCategories = [
    ...new Set(parseList(optionalEnv('MCP_DIRECT_TOOL_CATEGORIES')).map(value => value.toLowerCase()))
  ];
  const directTools = [
    ...new Set(parseList(optionalEnv('MCP_DIRECT_TOOLS')).map(value => value.toLowerCase()))
  ];
  const disabledGraphScopes = [
    ...new Set(parseList(optionalEnv('MCP_DISABLED_GRAPH_SCOPES')).map(scopeName))
  ];
  const graphResilience: GraphResilienceConfig = {
    maxRetries: positiveIntegerEnv('MCP_GRAPH_MAX_RETRIES', 3),
    initialBackoffMs: positiveIntegerEnv('MCP_GRAPH_INITIAL_BACKOFF_MS', 1000),
    backoffMultiplier: Math.max(1, positiveIntegerEnv('MCP_GRAPH_BACKOFF_MULTIPLIER', 2)),
    circuitBreakerThreshold: Math.max(1, positiveIntegerEnv('MCP_GRAPH_CIRCUIT_BREAKER_THRESHOLD', 5)),
    circuitBreakerCooldownMs: positiveIntegerEnv('MCP_GRAPH_CIRCUIT_BREAKER_COOLDOWN_MS', 30000),
    timeoutMs: Math.max(1, positiveIntegerEnv('MCP_GRAPH_TIMEOUT_MS', 30000))
  };

  const missing: string[] = [];
  if (!tenantId) missing.push('MS_TENANT_ID');
  if (!clientId) missing.push('MS_CLIENT_ID');

  const missingRemote: string[] = [];
  if (!tenantId) missingRemote.push('MS_TENANT_ID');
  if (!clientId) missingRemote.push('MS_CLIENT_ID');
  if (!clientSecret) missingRemote.push('MS_CLIENT_SECRET');
  if (oauthBridgeEnabled && !oauthBridgeMicrosoftClientId) {
    missingRemote.push(
      oauthBridgeMicrosoftClientType === 'confidential_web'
        ? 'MS_OAUTH_CLIENT_ID (or MS_CLIENT_ID)'
        : 'MS_OAUTH_CLIENT_ID (or MS_PUBLIC_CLIENT_ID)'
    );
  }
  if (
    oauthBridgeEnabled &&
    oauthBridgeMicrosoftClientType === 'confidential_web' &&
    !oauthBridgeMicrosoftClientSecret
  ) {
    missingRemote.push('MS_OAUTH_CLIENT_SECRET (or matching MS_CLIENT_SECRET)');
  }
  if (!inboundAuthDisabled && inboundTokenAudiences.length === 0) missingRemote.push('MCP_TOKEN_AUDIENCE');

  return {
    tenantId,
    clientId,
    clientSecret,
    publicClientId,
    oauthBridgeMicrosoftClientType,
    oauthBridgeMicrosoftClientId,
    oauthBridgeMicrosoftClientSecret,
    authorityHost,
    graphBaseUrl,
    graphResource,
    scopes: parseScopes(optionalEnv('MS_GRAPH_SCOPES'), graphResource),
    oboScopes: parseOboScopes(optionalEnv('MS_OBO_GRAPH_SCOPES'), graphResource),
    tokenCachePath,
    deviceCodeCachePath,
    enableRawGraphGet: booleanEnv('MS_ENABLE_RAW_GRAPH_GET', false),
    transport,
    httpPort,
    httpHost,
    httpPath,
    oauthHttpPath,
    httpStateless,
    publicBaseUrl,
    publicPathPrefix,
    resourceUrl,
    resourceMetadataUrl,
    authorizationServers,
    authorizationScopes,
    requiredTokenScopes,
    resourceDocumentationUrl: optionalEnv('MCP_RESOURCE_DOCUMENTATION_URL'),
    inboundAuthDisabled,
    allowUnauthenticatedDiscovery,
    inboundTokenAudiences,
    inboundTokenIssuer: optionalEnv('MCP_TOKEN_ISSUER'),
    oidcMetadataUrl,
    oauthBridgeEnabled,
    oauthBridgeIssuer,
    oauthBridgeRedirectUri,
    oauthBridgeScopes,
    oauthBridgePrompt,
    oauthBridgeCallbackDelivery,
    oauthBridgeStatePath,
    toolExposureMode,
    toolCategories,
    directToolCategories,
    directTools,
    disabledGraphScopes,
    roleBasedFiltering: booleanEnv('MCP_ROLE_BASED_FILTERING', true),
    auditLogEnabled: booleanEnv('MCP_AUDIT_LOG_ENABLED', true),
    auditLogPath: resolvePath(optionalEnv('MCP_AUDIT_LOG_PATH'), '.tokens/audit'),
    sendMode: parseSendMode(),
    confirmOperations: parseList(optionalEnv('MCP_CONFIRM_OPERATIONS')).map(value => value.toLowerCase()),
    confirmTtlSeconds: Math.max(1, positiveIntegerEnv('MCP_CONFIRM_TTL_SECONDS', 300)),
    graphResilience,
    adminToken: optionalEnv('MCP_ADMIN_TOKEN'),
    downstreamServices: parseDownstreamServices(optionalEnv('MCP_DOWNSTREAM_SERVICES')),
    missing,
    missingRemote
  };
}

export function requireConfigured(config: AppConfig): asserts config is AppConfig & {
  tenantId: string;
  clientId: string;
} {
  if (config.missing.length > 0) {
    throw new ConfigError(`Missing required environment variables: ${config.missing.join(', ')}`);
  }
}

export function requireOboConfigured(config: AppConfig): asserts config is AppConfig & {
  tenantId: string;
  clientId: string;
  clientSecret: string;
} {
  const missing = ['MS_TENANT_ID', 'MS_CLIENT_ID', 'MS_CLIENT_SECRET'].filter(name => {
    if (name === 'MS_TENANT_ID') return !config.tenantId;
    if (name === 'MS_CLIENT_ID') return !config.clientId;
    return !config.clientSecret;
  });

  if (missing.length > 0) {
    throw new ConfigError(`Missing required OBO environment variables: ${missing.join(', ')}`);
  }
}

export function requireRemoteConfigured(config: AppConfig): void {
  if (config.missingRemote.length > 0) {
    throw new ConfigError(`Missing required remote MCP environment variables: ${config.missingRemote.join(', ')}`);
  }
}
