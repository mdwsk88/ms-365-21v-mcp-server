import { createHash, randomBytes } from 'node:crypto';
import fs from 'node:fs/promises';
import path from 'node:path';
import { decodeJwt, type JWTPayload } from 'jose';
import type { AppConfig } from './config.js';
import { ConfigError } from './config.js';
import { mountedRoutePaths } from './http-routes.js';
import type { TokenCache } from './oauth.js';

type DynamicClient = {
  clientId: string;
  clientName?: string;
  redirectUris: string[];
  grantTypes: string[];
  responseTypes: string[];
  tokenEndpointAuthMethod: string;
  createdAt: number;
};

type PendingMicrosoftAuth = {
  clientId: string;
  redirectUri: string;
  scope: string[];
  clientState?: string;
  clientCodeChallenge?: string;
  clientCodeChallengeMethod?: string;
  microsoftCodeVerifier: string;
  expiresAt: number;
};

type BridgeAuthCode = {
  clientId: string;
  redirectUri: string;
  scope: string[];
  clientCodeChallenge?: string;
  clientCodeChallengeMethod?: string;
  microsoftToken: TokenCache;
  expiresAt: number;
};

type BridgeSession = {
  clientId: string;
  accessToken: string;
  refreshToken: string;
  scope: string[];
  microsoftToken: TokenCache;
  roles?: string[];
  expiresAt: number;
  createdAt: number;
};

type PersistedBridgeState = {
  clients: DynamicClient[];
  sessions: BridgeSession[];
};

type TokenResponse = {
  access_token: string;
  refresh_token?: string;
  id_token?: string;
  token_type: string;
  scope?: string;
  expires_in: number;
};

type OAuthErrorBody = {
  error?: string;
  error_description?: string;
};

type BridgeAuthResult = {
  userAssertion: string;
  claims: JWTPayload;
};

const pendingMicrosoftAuth = new Map<string, PendingMicrosoftAuth>();
const bridgeAuthCodes = new Map<string, BridgeAuthCode>();
const stateCache = new Map<string, PersistedBridgeState>();
const stateLoadPromises = new Map<string, Promise<PersistedBridgeState>>();
const stateWriteQueues = new Map<string, Promise<void>>();

function oauthBridgeLog(event: string, details: Record<string, unknown> = {}): void {
  console.error(JSON.stringify({ component: 'oauth_bridge', event, ...details }));
}

function safeUrlDetails(value: string | undefined): Record<string, string> {
  if (!value) return {};
  try {
    const url = new URL(value);
    return { urlOrigin: url.origin, urlPath: url.pathname };
  } catch {
    return { url: 'invalid-url' };
  }
}

function nowSeconds(): number {
  return Math.floor(Date.now() / 1000);
}

function randomToken(prefix: string): string {
  return `${prefix}_${randomBytes(32).toString('base64url')}`;
}

function base64UrlSha256(value: string): string {
  return createHash('sha256').update(value).digest('base64url');
}

function toTokenCache(response: TokenResponse): TokenCache {
  const acquiredAt = nowSeconds();
  return {
    accessToken: response.access_token,
    refreshToken: response.refresh_token,
    idToken: response.id_token,
    tokenType: response.token_type,
    scopes: response.scope,
    expiresAt: acquiredAt + response.expires_in,
    acquiredAt
  };
}

function decodeTokenClaims(token: string | undefined): JWTPayload | undefined {
  if (!token) return undefined;
  try {
    return decodeJwt(token);
  } catch {
    return undefined;
  }
}

function claimStrings(value: unknown): string[] {
  if (Array.isArray(value)) {
    return value.filter((item): item is string => typeof item === 'string');
  }
  return typeof value === 'string' ? value.split(/\s+/).filter(Boolean) : [];
}

function microsoftTokenRoles(token: TokenCache): string[] {
  // API authorization roles belong to the resource access token, not the client ID token.
  const accessClaims = decodeTokenClaims(token.accessToken);
  return [...new Set(claimStrings(accessClaims?.roles))];
}

function microsoftEndpoint(config: AppConfig, suffix: 'authorize' | 'token'): string {
  if (!config.tenantId) {
    throw new ConfigError('MS_TENANT_ID is required for the OAuth bridge.');
  }
  return `${config.authorityHost}/${encodeURIComponent(config.tenantId)}/oauth2/v2.0/${suffix}`;
}

function microsoftTokenRequest(
  config: AppConfig,
  params: Record<string, string>
): Record<string, string> {
  if (!config.oauthBridgeMicrosoftClientId) {
    throw new ConfigError('A Microsoft OAuth client ID is required for the OAuth bridge.');
  }

  const request: Record<string, string> = {
    ...params,
    client_id: config.oauthBridgeMicrosoftClientId
  };
  if (config.oauthBridgeMicrosoftClientType === 'confidential_web') {
    if (!config.oauthBridgeMicrosoftClientSecret) {
      throw new ConfigError('A Microsoft OAuth client secret is required for confidential_web mode.');
    }
    request.client_secret = config.oauthBridgeMicrosoftClientSecret;
  }
  return request;
}

function localUrl(config: AppConfig, pathname: string): string {
  return `${config.oauthBridgeIssuer}${pathname}`;
}

function splitScope(scope: string | undefined, fallback: string[]): string[] {
  const parsed = scope?.split(/\s+/).map(item => item.trim()).filter(Boolean) ?? [];
  return parsed.length ? parsed : fallback;
}

function finalScopeSegment(scope: string): string {
  const trimmed = scope.replace(/\/+$/, '');
  const slashIndex = trimmed.lastIndexOf('/');
  return slashIndex >= 0 ? trimmed.slice(slashIndex + 1) : trimmed;
}

function includesScope(tokenScopes: string[], requiredScope: string): boolean {
  const scopeSet = new Set(tokenScopes);
  return scopeSet.has(requiredScope) || scopeSet.has(finalScopeSegment(requiredScope));
}

function bridgeSessionScopes(config: AppConfig, requestedScopes: string[]): string[] {
  const scopes = new Set<string>();
  for (const scope of requestedScopes) {
    scopes.add(scope);
  }
  for (const scope of config.requiredTokenScopes) {
    scopes.add(scope);
  }
  for (const scope of config.authorizationScopes) {
    scopes.add(finalScopeSegment(scope));
  }
  return [...scopes];
}

function pruneExpiredEphemeralState(): void {
  const now = nowSeconds();
  for (const [state, value] of pendingMicrosoftAuth) {
    if (value.expiresAt <= now) pendingMicrosoftAuth.delete(state);
  }
  for (const [code, value] of bridgeAuthCodes) {
    if (value.expiresAt <= now) bridgeAuthCodes.delete(code);
  }
}

async function loadBridgeState(config: AppConfig): Promise<PersistedBridgeState> {
  const cached = stateCache.get(config.oauthBridgeStatePath);
  if (cached) return cached;

  const existingLoad = stateLoadPromises.get(config.oauthBridgeStatePath);
  if (existingLoad) return existingLoad;

  const load = (async () => {
    try {
      const parsed = JSON.parse(await fs.readFile(config.oauthBridgeStatePath, 'utf8')) as PersistedBridgeState;
      const state = {
        clients: Array.isArray(parsed.clients) ? parsed.clients : [],
        sessions: Array.isArray(parsed.sessions) ? parsed.sessions : []
      };
      stateCache.set(config.oauthBridgeStatePath, state);
      return state;
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error;
      const state = { clients: [], sessions: [] };
      stateCache.set(config.oauthBridgeStatePath, state);
      return state;
    }
  })();
  stateLoadPromises.set(config.oauthBridgeStatePath, load);
  try {
    return await load;
  } finally {
    if (stateLoadPromises.get(config.oauthBridgeStatePath) === load) {
      stateLoadPromises.delete(config.oauthBridgeStatePath);
    }
  }
}

async function saveBridgeState(config: AppConfig, state: PersistedBridgeState): Promise<void> {
  const statePath = config.oauthBridgeStatePath;
  const previousWrite = stateWriteQueues.get(statePath) ?? Promise.resolve();
  const write = previousWrite.catch(() => undefined).then(async () => {
    await fs.mkdir(path.dirname(statePath), { recursive: true });
    const tempPath = `${statePath}.${process.pid}.${randomBytes(8).toString('hex')}.tmp`;
    try {
      await fs.writeFile(tempPath, `${JSON.stringify(state, null, 2)}\n`, { mode: 0o600 });
      await fs.chmod(tempPath, 0o600);
      await fs.rename(tempPath, statePath);
    } finally {
      await fs.rm(tempPath, { force: true }).catch(() => undefined);
    }
  });
  stateWriteQueues.set(statePath, write);
  try {
    await write;
  } finally {
    if (stateWriteQueues.get(statePath) === write) stateWriteQueues.delete(statePath);
  }
}

function formBody(params: Record<string, string>): URLSearchParams {
  const body = new URLSearchParams();
  for (const [key, value] of Object.entries(params)) {
    body.set(key, value);
  }
  return body;
}

async function postMicrosoftToken(config: AppConfig, params: Record<string, string>): Promise<TokenCache> {
  const response = await fetch(microsoftEndpoint(config, 'token'), {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded', Accept: 'application/json' },
    body: formBody(params)
  });
  const text = await response.text();
  const body = text ? (JSON.parse(text) as TokenResponse & OAuthErrorBody) : ({} as TokenResponse & OAuthErrorBody);

  if (!response.ok || body.error) {
    const message = body.error_description ?? body.error ?? `Microsoft token request failed with HTTP ${response.status}`;
    const error = new Error(message);
    error.name = body.error ?? 'oauth_error';
    throw error;
  }

  return toTokenCache(body);
}

function queryParam(value: unknown): string | undefined {
  if (Array.isArray(value)) return queryParam(value[0]);
  return typeof value === 'string' ? value : undefined;
}

async function requestFormParams(req: any): Promise<Record<string, string>> {
  if (req.body && typeof req.body === 'object' && Object.keys(req.body).length > 0) {
    return Object.fromEntries(
      Object.entries(req.body).flatMap(([key, value]) => (typeof value === 'string' ? [[key, value]] : []))
    );
  }

  const chunks: Buffer[] = [];
  for await (const chunk of req) {
    chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
  }
  return Object.fromEntries(new URLSearchParams(Buffer.concat(chunks).toString('utf8')));
}

function sendOAuthError(res: any, status: number, error: string, description: string): void {
  res.status(status).json({ error, error_description: description });
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function isLoopbackRedirect(value: string): boolean {
  try {
    const url = new URL(value);
    const hostname = url.hostname.toLowerCase();
    return hostname === 'localhost' || hostname === '127.0.0.1' || hostname === '[::1]' || hostname === '::1';
  } catch {
    return false;
  }
}

function callbackDeliveryPage(callbackUrl: string): string {
  const callbackUrlJson = JSON.stringify(callbackUrl);
  const callbackUrlHtml = escapeHtml(callbackUrl);

  return `<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>授权完成</title>
  <style>
    :root { color-scheme: light; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; }
    body { margin: 0; min-height: 100vh; display: grid; place-items: center; background: #f7f8fa; color: #1f2328; }
    main { max-width: 420px; padding: 32px; text-align: center; }
    h1 { margin: 0 0 12px; font-size: 22px; font-weight: 650; }
    p { margin: 0 0 20px; color: #59636e; line-height: 1.6; }
    a { color: #0969da; }
  </style>
</head>
<body>
  <main>
    <h1>授权完成</h1>
    <p id="status">正在返回 MCP 客户端，可以关闭此页面。</p>
    <p><a id="fallback" href="${callbackUrlHtml}" rel="noreferrer">如果没有自动完成，点击这里继续</a></p>
  </main>
  <iframe id="callback-frame" title="OAuth callback" referrerpolicy="no-referrer" style="display:none"></iframe>
  <script>
    const callbackUrl = ${callbackUrlJson};
    const statusEl = document.getElementById('status');
    let finished = false;

    function finish() {
      if (finished) return;
      finished = true;
      statusEl.textContent = '已返回 MCP 客户端，可以关闭此页面。';
      setTimeout(() => window.close(), 600);
    }

    async function deliver() {
      try {
        await fetch(callbackUrl, {
          method: 'GET',
          mode: 'no-cors',
          credentials: 'include',
          cache: 'no-store',
          referrerPolicy: 'no-referrer'
        });
        finish();
        return;
      } catch (_error) {
        const frame = document.getElementById('callback-frame');
        frame.addEventListener('load', finish, { once: true });
        frame.src = callbackUrl;
        setTimeout(finish, 2500);
      }
    }

    deliver();
  </script>
</body>
</html>`;
}

function createBridgeClaims(config: AppConfig, session: BridgeSession): JWTPayload {
  const accessClaims = decodeTokenClaims(session.microsoftToken.accessToken);
  const idClaims = decodeTokenClaims(session.microsoftToken.idToken);
  return {
    iss: config.oauthBridgeIssuer,
    aud: config.inboundTokenAudiences[0] ?? config.resourceUrl,
    sub: accessClaims?.sub ?? idClaims?.sub ?? session.clientId,
    azp: session.clientId,
    oid: accessClaims?.oid ?? idClaims?.oid,
    name: accessClaims?.name ?? idClaims?.name,
    preferred_username: accessClaims?.preferred_username ?? idClaims?.preferred_username,
    upn: accessClaims?.upn ?? idClaims?.upn,
    roles: session.roles ?? microsoftTokenRoles(session.microsoftToken),
    scp: session.scope.join(' '),
    iat: session.createdAt,
    exp: session.expiresAt
  };
}

function findClient(state: PersistedBridgeState, clientId: string | undefined): DynamicClient | undefined {
  return clientId ? state.clients.find(client => client.clientId === clientId) : undefined;
}

function validatePkce(codeChallenge: string | undefined, method: string | undefined, verifier: string | undefined): boolean {
  if (!codeChallenge) return true;
  if (!verifier) return false;
  if (method && method !== 'S256') return false;
  return base64UrlSha256(verifier) === codeChallenge;
}

async function refreshMicrosoftTokenForSession(config: AppConfig, session: BridgeSession): Promise<boolean> {
  if (session.microsoftToken.expiresAt > nowSeconds() + 300) return true;
  if (!session.microsoftToken.refreshToken || !config.oauthBridgeMicrosoftClientId) return false;
  if (
    config.oauthBridgeMicrosoftClientType === 'confidential_web' &&
    !config.oauthBridgeMicrosoftClientSecret
  ) {
    return false;
  }

  oauthBridgeLog('microsoft_refresh_token_start', {
    clientId: session.clientId,
    microsoftClientType: config.oauthBridgeMicrosoftClientType
  });
  session.microsoftToken = await postMicrosoftToken(
    config,
    microsoftTokenRequest(config, {
      grant_type: 'refresh_token',
      refresh_token: session.microsoftToken.refreshToken,
      scope: config.oauthBridgeScopes.join(' ')
    })
  );
  session.roles = microsoftTokenRoles(session.microsoftToken);
  session.expiresAt = Math.min(session.microsoftToken.expiresAt, nowSeconds() + 3600);
  oauthBridgeLog('microsoft_refresh_token_success', { clientId: session.clientId });
  return true;
}

async function createBridgeSession(
  config: AppConfig,
  state: PersistedBridgeState,
  authCode: BridgeAuthCode
): Promise<BridgeSession> {
  const session: BridgeSession = {
    clientId: authCode.clientId,
    accessToken: randomToken('mcp_at'),
    refreshToken: randomToken('mcp_rt'),
    scope: bridgeSessionScopes(config, authCode.scope),
    microsoftToken: authCode.microsoftToken,
    roles: microsoftTokenRoles(authCode.microsoftToken),
    expiresAt: Math.min(authCode.microsoftToken.expiresAt, nowSeconds() + 3600),
    createdAt: nowSeconds()
  };

  state.sessions = state.sessions.filter(existing => existing.expiresAt > nowSeconds());
  state.sessions.push(session);
  await saveBridgeState(config, state);
  return session;
}

export function oauthBridgeAuthorizationServerMetadata(config: AppConfig) {
  return {
    issuer: config.oauthBridgeIssuer,
    authorization_endpoint: localUrl(config, '/oauth/authorize'),
    token_endpoint: localUrl(config, '/oauth/token'),
    registration_endpoint: localUrl(config, '/oauth/register'),
    response_types_supported: ['code'],
    grant_types_supported: ['authorization_code', 'refresh_token'],
    token_endpoint_auth_methods_supported: ['none'],
    code_challenge_methods_supported: ['S256'],
    scopes_supported: config.authorizationScopes,
    service_documentation: config.resourceDocumentationUrl
  };
}

function oauthBridgeOpenIdConfigurationMetadata(config: AppConfig) {
  return {
    ...oauthBridgeAuthorizationServerMetadata(config),
    // AgentRun mounts the app below /tools/<name>. Path-aware MCP clients can
    // only reach the appended OIDC fallback, so keep it schema-compatible.
    jwks_uri: localUrl(config, '/.well-known/jwks.json'),
    subject_types_supported: ['public'],
    id_token_signing_alg_values_supported: ['RS256']
  };
}

export function mountOAuthBridgeRoutes(app: any, config: AppConfig): void {
  if (!config.oauthBridgeEnabled) return;

  const mountGet = (pathname: string, handler: (req: any, res: any) => void | Promise<void>) => {
    for (const routePath of mountedRoutePaths(config, pathname)) app.get(routePath, handler);
  };
  const mountPost = (pathname: string, handler: (req: any, res: any) => void | Promise<void>) => {
    for (const routePath of mountedRoutePaths(config, pathname)) app.post(routePath, handler);
  };

  mountGet('/.well-known/oauth-authorization-server', (_req: any, res: any) => {
    res.json(oauthBridgeAuthorizationServerMetadata(config));
  });
  mountGet('/.well-known/openid-configuration', (_req: any, res: any) => {
    res.json(oauthBridgeOpenIdConfigurationMetadata(config));
  });
  mountGet('/.well-known/jwks.json', (_req: any, res: any) => {
    res.json({ keys: [] });
  });

  mountPost('/oauth/register', async (req: any, res: any) => {
    const body = req.body && typeof req.body === 'object' ? req.body : {};
    const redirectUris = Array.isArray(body.redirect_uris)
      ? body.redirect_uris.filter((value: unknown): value is string => typeof value === 'string')
      : [];

    if (redirectUris.length === 0) {
      sendOAuthError(res, 400, 'invalid_client_metadata', 'redirect_uris is required.');
      return;
    }

    const client: DynamicClient = {
      clientId: randomToken('mcp_client'),
      clientName: typeof body.client_name === 'string' ? body.client_name : 'MCP OAuth Client',
      redirectUris,
      grantTypes: ['authorization_code', 'refresh_token'],
      responseTypes: ['code'],
      tokenEndpointAuthMethod: 'none',
      createdAt: nowSeconds()
    };

    const state = await loadBridgeState(config);
    state.clients.push(client);
    await saveBridgeState(config, state);

    oauthBridgeLog('dynamic_client_registered', {
      clientName: client.clientName,
      redirectUriCount: client.redirectUris.length,
      firstRedirect: safeUrlDetails(client.redirectUris[0])
    });

    res.status(201).json({
      client_id: client.clientId,
      client_id_issued_at: client.createdAt,
      client_name: client.clientName,
      redirect_uris: client.redirectUris,
      grant_types: client.grantTypes,
      response_types: client.responseTypes,
      token_endpoint_auth_method: client.tokenEndpointAuthMethod
    });
  });

  mountGet('/oauth/authorize', async (req: any, res: any) => {
    try {
      pruneExpiredEphemeralState();
      if (!config.oauthBridgeMicrosoftClientId) {
        sendOAuthError(res, 500, 'server_error', 'A Microsoft OAuth client ID is required for the OAuth bridge.');
        return;
      }
      if (
        config.oauthBridgeMicrosoftClientType === 'confidential_web' &&
        !config.oauthBridgeMicrosoftClientSecret
      ) {
        sendOAuthError(res, 500, 'server_error', 'A Microsoft OAuth client secret is required for confidential_web mode.');
        return;
      }

      const responseType = queryParam(req.query.response_type);
      const clientId = queryParam(req.query.client_id);
      const redirectUri = queryParam(req.query.redirect_uri);
      const clientState = queryParam(req.query.state);
      const codeChallenge = queryParam(req.query.code_challenge);
      const codeChallengeMethod = queryParam(req.query.code_challenge_method);
      const scope = splitScope(queryParam(req.query.scope), config.authorizationScopes);

      if (responseType !== 'code' || !clientId || !redirectUri) {
        sendOAuthError(res, 400, 'invalid_request', 'response_type=code, client_id, and redirect_uri are required.');
        return;
      }
      if (codeChallengeMethod && codeChallengeMethod !== 'S256') {
        sendOAuthError(res, 400, 'invalid_request', 'Only S256 PKCE is supported.');
        return;
      }

      const state = await loadBridgeState(config);
      const client = findClient(state, clientId);
      if (!client || !client.redirectUris.includes(redirectUri)) {
        oauthBridgeLog('authorize_rejected_unknown_client_or_redirect', {
          clientKnown: Boolean(client),
          ...safeUrlDetails(redirectUri)
        });
        sendOAuthError(res, 400, 'invalid_request', 'Unknown client_id or redirect_uri.');
        return;
      }

      const microsoftState = randomToken('ms_state');
      const microsoftCodeVerifier = randomBytes(32).toString('base64url');
      pendingMicrosoftAuth.set(microsoftState, {
        clientId,
        redirectUri,
        scope,
        clientState,
        clientCodeChallenge: codeChallenge,
        clientCodeChallengeMethod: codeChallengeMethod,
        microsoftCodeVerifier,
        expiresAt: nowSeconds() + 10 * 60
      });

      const authorizeUrl = new URL(microsoftEndpoint(config, 'authorize'));
      authorizeUrl.searchParams.set('client_id', config.oauthBridgeMicrosoftClientId);
      authorizeUrl.searchParams.set('response_type', 'code');
      authorizeUrl.searchParams.set('redirect_uri', config.oauthBridgeRedirectUri);
      authorizeUrl.searchParams.set('response_mode', 'query');
      authorizeUrl.searchParams.set('scope', config.oauthBridgeScopes.join(' '));
      authorizeUrl.searchParams.set('state', microsoftState);
      authorizeUrl.searchParams.set('code_challenge', base64UrlSha256(microsoftCodeVerifier));
      authorizeUrl.searchParams.set('code_challenge_method', 'S256');
      if (config.oauthBridgePrompt) {
        authorizeUrl.searchParams.set('prompt', config.oauthBridgePrompt);
      }

      oauthBridgeLog('authorize_redirect_to_microsoft', {
        clientRedirect: safeUrlDetails(redirectUri),
        microsoftRedirect: safeUrlDetails(config.oauthBridgeRedirectUri),
        microsoftClientType: config.oauthBridgeMicrosoftClientType,
        prompt: config.oauthBridgePrompt,
        scopeCount: scope.length
      });
      res.redirect(authorizeUrl.toString());
    } catch (error) {
      oauthBridgeLog('authorize_error', { error: error instanceof Error ? error.message : String(error) });
      sendOAuthError(res, 500, 'server_error', error instanceof Error ? error.message : String(error));
    }
  });

  mountGet('/oauth/microsoft/callback', async (req: any, res: any) => {
    try {
      pruneExpiredEphemeralState();
      const stateParam = queryParam(req.query.state);
      const code = queryParam(req.query.code);
      const microsoftError = queryParam(req.query.error);
      const pending = stateParam ? pendingMicrosoftAuth.get(stateParam) : undefined;

      if (!pending || !stateParam) {
        oauthBridgeLog('microsoft_callback_rejected_unknown_state', { hasState: Boolean(stateParam) });
        sendOAuthError(res, 400, 'invalid_request', 'Unknown or expired OAuth state.');
        return;
      }
      pendingMicrosoftAuth.delete(stateParam);

      if (microsoftError) {
        oauthBridgeLog('microsoft_callback_error', {
          error: microsoftError,
          clientRedirect: safeUrlDetails(pending.redirectUri)
        });
        const redirectUrl = new URL(pending.redirectUri);
        redirectUrl.searchParams.set('error', microsoftError);
        const description = queryParam(req.query.error_description);
        if (description) redirectUrl.searchParams.set('error_description', description);
        if (pending.clientState) redirectUrl.searchParams.set('state', pending.clientState);
        res.redirect(redirectUrl.toString());
        return;
      }
      if (!code || !config.oauthBridgeMicrosoftClientId) {
        oauthBridgeLog('microsoft_callback_missing_code', { hasCode: Boolean(code) });
        sendOAuthError(res, 400, 'invalid_request', 'Missing Microsoft authorization code.');
        return;
      }

      oauthBridgeLog('microsoft_token_exchange_start', {
        microsoftRedirect: safeUrlDetails(config.oauthBridgeRedirectUri),
        microsoftClientType: config.oauthBridgeMicrosoftClientType
      });
      const microsoftToken = await postMicrosoftToken(
        config,
        microsoftTokenRequest(config, {
          grant_type: 'authorization_code',
          code,
          redirect_uri: config.oauthBridgeRedirectUri,
          code_verifier: pending.microsoftCodeVerifier,
          scope: config.oauthBridgeScopes.join(' ')
        })
      );

      const bridgeCode = randomToken('mcp_code');
      bridgeAuthCodes.set(bridgeCode, {
        clientId: pending.clientId,
        redirectUri: pending.redirectUri,
        scope: pending.scope,
        clientCodeChallenge: pending.clientCodeChallenge,
        clientCodeChallengeMethod: pending.clientCodeChallengeMethod,
        microsoftToken,
        expiresAt: nowSeconds() + 5 * 60
      });

      const redirectUrl = new URL(pending.redirectUri);
      redirectUrl.searchParams.set('code', bridgeCode);
      if (pending.clientState) redirectUrl.searchParams.set('state', pending.clientState);
      const callbackUrl = redirectUrl.toString();
      const backgroundDelivery =
        config.oauthBridgeCallbackDelivery === 'background' && isLoopbackRedirect(callbackUrl);
      oauthBridgeLog(backgroundDelivery ? 'deliver_callback_in_background' : 'redirect_back_to_client', {
        clientRedirect: safeUrlDetails(pending.redirectUri)
      });
      if (backgroundDelivery) {
        res.status(200).type('html').send(callbackDeliveryPage(callbackUrl));
      } else {
        res.redirect(callbackUrl);
      }
    } catch (error) {
      oauthBridgeLog('microsoft_callback_error_unhandled', {
        error: error instanceof Error ? error.message : String(error)
      });
      sendOAuthError(res, 500, 'server_error', error instanceof Error ? error.message : String(error));
    }
  });

  mountPost('/oauth/token', async (req: any, res: any) => {
    try {
      pruneExpiredEphemeralState();
      const params = await requestFormParams(req);
      const grantType = params.grant_type;
      const state = await loadBridgeState(config);
      oauthBridgeLog('token_request', {
        grantType,
        hasClientId: Boolean(params.client_id),
        hasRedirectUri: Boolean(params.redirect_uri),
        hasCodeVerifier: Boolean(params.code_verifier)
      });

      res.set('Cache-Control', 'no-store');
      res.set('Pragma', 'no-cache');

      if (grantType === 'authorization_code') {
        const code = params.code;
        const authCode = code ? bridgeAuthCodes.get(code) : undefined;
        const client = findClient(state, params.client_id);
        if (!authCode || !client || authCode.clientId !== client.clientId) {
          oauthBridgeLog('token_rejected_invalid_grant', {
            hasAuthCode: Boolean(authCode),
            clientKnown: Boolean(client)
          });
          sendOAuthError(res, 400, 'invalid_grant', 'Unknown authorization code or client_id.');
          return;
        }
        if (params.redirect_uri !== authCode.redirectUri) {
          oauthBridgeLog('token_rejected_redirect_mismatch', {
            requestRedirect: safeUrlDetails(params.redirect_uri),
            expectedRedirect: safeUrlDetails(authCode.redirectUri)
          });
          sendOAuthError(res, 400, 'invalid_grant', 'redirect_uri does not match the authorization request.');
          return;
        }
        if (!validatePkce(authCode.clientCodeChallenge, authCode.clientCodeChallengeMethod, params.code_verifier)) {
          oauthBridgeLog('token_rejected_pkce');
          sendOAuthError(res, 400, 'invalid_grant', 'PKCE verification failed.');
          return;
        }

        bridgeAuthCodes.delete(code);
        const session = await createBridgeSession(config, state, authCode);
        oauthBridgeLog('token_granted_authorization_code', {
          expiresIn: Math.max(60, session.expiresAt - nowSeconds()),
          scopeCount: session.scope.length
        });
        res.json({
          access_token: session.accessToken,
          refresh_token: session.refreshToken,
          token_type: 'Bearer',
          expires_in: Math.max(60, session.expiresAt - nowSeconds()),
          scope: session.scope.join(' ')
        });
        return;
      }

      if (grantType === 'refresh_token') {
        const session = state.sessions.find(existing => existing.refreshToken === params.refresh_token);
        if (!session) {
          oauthBridgeLog('token_rejected_unknown_refresh_token');
          sendOAuthError(res, 400, 'invalid_grant', 'Unknown refresh token.');
          return;
        }
        await refreshMicrosoftTokenForSession(config, session);
        session.scope = bridgeSessionScopes(config, session.scope);
        session.roles = session.roles ?? microsoftTokenRoles(session.microsoftToken);
        session.accessToken = randomToken('mcp_at');
        session.expiresAt = Math.min(session.microsoftToken.expiresAt, nowSeconds() + 3600);
        await saveBridgeState(config, state);
        oauthBridgeLog('token_granted_refresh_token', {
          expiresIn: Math.max(60, session.expiresAt - nowSeconds()),
          scopeCount: session.scope.length
        });
        res.json({
          access_token: session.accessToken,
          refresh_token: session.refreshToken,
          token_type: 'Bearer',
          expires_in: Math.max(60, session.expiresAt - nowSeconds()),
          scope: session.scope.join(' ')
        });
        return;
      }

      oauthBridgeLog('token_rejected_unsupported_grant', { grantType });
      sendOAuthError(res, 400, 'unsupported_grant_type', 'Only authorization_code and refresh_token are supported.');
    } catch (error) {
      oauthBridgeLog('token_error_unhandled', { error: error instanceof Error ? error.message : String(error) });
      sendOAuthError(res, 500, 'server_error', error instanceof Error ? error.message : String(error));
    }
  });
}

export async function authenticateOAuthBridgeToken(
  config: AppConfig,
  accessToken: string
): Promise<BridgeAuthResult | undefined> {
  if (!config.oauthBridgeEnabled) return undefined;

  const state = await loadBridgeState(config);
  const session = state.sessions.find(existing => existing.accessToken === accessToken);
  if (!session || session.expiresAt <= nowSeconds()) {
    oauthBridgeLog('bearer_rejected_unknown_or_expired_session', { hasSession: Boolean(session) });
    return undefined;
  }

  session.scope = bridgeSessionScopes(config, session.scope);
  session.roles = session.roles ?? microsoftTokenRoles(session.microsoftToken);
  const missingScopes = config.requiredTokenScopes.filter(scope => !includesScope(session.scope, scope));
  if (missingScopes.length > 0) {
    oauthBridgeLog('bearer_rejected_missing_scopes', { missingScopes });
    return undefined;
  }

  const refreshed = await refreshMicrosoftTokenForSession(config, session);
  if (!refreshed) {
    oauthBridgeLog('bearer_rejected_refresh_failed');
    return undefined;
  }
  await saveBridgeState(config, state);

  return {
    userAssertion: session.microsoftToken.accessToken,
    claims: createBridgeClaims(config, session)
  };
}
