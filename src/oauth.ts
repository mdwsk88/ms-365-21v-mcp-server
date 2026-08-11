import fs from 'node:fs/promises';
import path from 'node:path';
import { ConfidentialClientApplication, type AuthenticationResult } from '@azure/msal-node';
import type { AppConfig } from './config.js';
import { ConfigError, requireConfigured, requireOboConfigured } from './config.js';
import { getRequestContext } from './request-context.js';

export type DeviceCodeState = {
  deviceCode: string;
  userCode: string;
  verificationUri: string;
  verificationUriComplete?: string;
  expiresAt: number;
  intervalSeconds: number;
  message: string;
  scopes: string[];
  createdAt: number;
};

export type TokenCache = {
  accessToken: string;
  refreshToken?: string;
  idToken?: string;
  tokenType: string;
  scopes?: string;
  expiresAt: number;
  acquiredAt: number;
};

type OboCacheEntry = {
  token: TokenCache;
  assertion: string;
  scopes: string;
};

type DeviceCodeResponse = {
  device_code: string;
  user_code: string;
  verification_uri: string;
  verification_uri_complete?: string;
  expires_in: number;
  interval?: number;
  message: string;
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

export class OAuthProtocolError extends Error {
  readonly code?: string;
  readonly status?: number;

  constructor(message: string, code?: string, status?: number) {
    super(message);
    this.name = 'OAuthProtocolError';
    this.code = code;
    this.status = status;
  }
}

const oboCache = new Map<string, OboCacheEntry>();
const confidentialClients = new Map<string, ConfidentialClientApplication>();

function nowSeconds(): number {
  return Math.floor(Date.now() / 1000);
}

function endpoint(config: AppConfig, suffix: 'devicecode' | 'token'): string {
  requireConfigured(config);
  return `${config.authorityHost}/${encodeURIComponent(config.tenantId)}/oauth2/v2.0/${suffix}`;
}

function authority(config: AppConfig): string {
  requireConfigured(config);
  return `${config.authorityHost}/${encodeURIComponent(config.tenantId)}`;
}

function toForm(params: Record<string, string>): URLSearchParams {
  const body = new URLSearchParams();
  for (const [key, value] of Object.entries(params)) {
    body.set(key, value);
  }
  return body;
}

async function readJson<T>(filePath: string): Promise<T | undefined> {
  try {
    return JSON.parse(await fs.readFile(filePath, 'utf8')) as T;
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === 'ENOENT') return undefined;
    throw error;
  }
}

async function writeJsonSecure(filePath: string, value: unknown): Promise<void> {
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  const tempPath = `${filePath}.${process.pid}.tmp`;
  await fs.writeFile(tempPath, `${JSON.stringify(value, null, 2)}\n`, { mode: 0o600 });
  await fs.chmod(tempPath, 0o600);
  await fs.rename(tempPath, filePath);
}

async function removeIfExists(filePath: string): Promise<void> {
  try {
    await fs.rm(filePath, { force: true });
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error;
  }
}

async function postForm<T>(url: string, params: Record<string, string>): Promise<T> {
  const response = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: toForm(params)
  });
  const text = await response.text();
  const body = text ? (JSON.parse(text) as T & OAuthErrorBody) : ({} as T & OAuthErrorBody);

  if (!response.ok || body.error) {
    throw new OAuthProtocolError(
      body.error_description ?? body.error ?? `OAuth request failed with HTTP ${response.status}`,
      body.error,
      response.status
    );
  }

  return body;
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

function toMsalTokenCache(response: AuthenticationResult): TokenCache {
  if (!response.accessToken) {
    throw new OAuthProtocolError('MSAL OBO response did not include an access token.');
  }

  return {
    accessToken: response.accessToken,
    idToken: response.idToken,
    tokenType: response.tokenType || 'Bearer',
    scopes: response.scopes?.join(' '),
    expiresAt: Math.floor((response.expiresOn?.getTime() ?? Date.now() + 60 * 60 * 1000) / 1000),
    acquiredAt: nowSeconds()
  };
}

function getConfidentialClient(config: AppConfig): ConfidentialClientApplication {
  requireOboConfigured(config);
  const clientKey = `${config.authorityHost}:${config.tenantId}:${config.clientId}`;
  const cached = confidentialClients.get(clientKey);
  if (cached) return cached;

  const client = new ConfidentialClientApplication({
    auth: {
      clientId: config.clientId,
      clientSecret: config.clientSecret,
      authority: authority(config),
      knownAuthorities: [new URL(config.authorityHost).host]
    }
  });

  confidentialClients.set(clientKey, client);
  return client;
}

function oboCacheKey(config: AppConfig, assertion: string, scopes: string[]): string {
  return `${config.clientId ?? ''}:${scopes.join(' ')}:${assertion}`;
}

export async function startDeviceLogin(config: AppConfig, scopes = config.scopes): Promise<DeviceCodeState> {
  requireConfigured(config);
  const response = await postForm<DeviceCodeResponse>(endpoint(config, 'devicecode'), {
    client_id: config.clientId,
    scope: scopes.join(' ')
  });

  const state: DeviceCodeState = {
    deviceCode: response.device_code,
    userCode: response.user_code,
    verificationUri: response.verification_uri,
    verificationUriComplete: response.verification_uri_complete,
    expiresAt: nowSeconds() + response.expires_in,
    intervalSeconds: response.interval ?? 5,
    message: response.message,
    scopes,
    createdAt: nowSeconds()
  };

  await writeJsonSecure(config.deviceCodeCachePath, state);
  return state;
}

function sleep(ms: number): Promise<void> {
  return new Promise(resolve => setTimeout(resolve, ms));
}

export async function completeDeviceLogin(config: AppConfig, timeoutSeconds = 120): Promise<TokenCache> {
  requireConfigured(config);
  const state = await readJson<DeviceCodeState>(config.deviceCodeCachePath);
  if (!state) {
    throw new ConfigError('No pending device login. Call auth_start_device_login first.');
  }

  const deadline = Math.min(nowSeconds() + timeoutSeconds, state.expiresAt);
  let intervalSeconds = state.intervalSeconds;

  while (nowSeconds() <= deadline) {
    try {
      const response = await postForm<TokenResponse>(endpoint(config, 'token'), {
        client_id: config.clientId,
        grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
        device_code: state.deviceCode
      });
      const token = toTokenCache(response);
      await writeJsonSecure(config.tokenCachePath, token);
      await removeIfExists(config.deviceCodeCachePath);
      return token;
    } catch (error) {
      if (!(error instanceof OAuthProtocolError)) throw error;

      if (error.code === 'authorization_pending') {
        await sleep(intervalSeconds * 1000);
        continue;
      }
      if (error.code === 'slow_down') {
        intervalSeconds += 5;
        await sleep(intervalSeconds * 1000);
        continue;
      }
      if (error.code === 'expired_token') {
        await removeIfExists(config.deviceCodeCachePath);
      }
      throw error;
    }
  }

  throw new OAuthProtocolError(
    'Device login is still pending. Complete browser sign-in and call this tool again.',
    'authorization_pending'
  );
}

export async function getTokenCache(config: AppConfig): Promise<TokenCache | undefined> {
  return readJson<TokenCache>(config.tokenCachePath);
}

export async function getPendingDeviceLogin(config: AppConfig): Promise<DeviceCodeState | undefined> {
  const state = await readJson<DeviceCodeState>(config.deviceCodeCachePath);
  if (state && state.expiresAt <= nowSeconds()) {
    await removeIfExists(config.deviceCodeCachePath);
    return undefined;
  }
  return state;
}

async function refreshAccessToken(config: AppConfig, token: TokenCache): Promise<TokenCache> {
  requireConfigured(config);
  if (!token.refreshToken) {
    throw new ConfigError('No refresh token is cached. Call auth_start_device_login first.');
  }

  const response = await postForm<TokenResponse>(endpoint(config, 'token'), {
    client_id: config.clientId,
    grant_type: 'refresh_token',
    refresh_token: token.refreshToken,
    scope: config.scopes.join(' ')
  });
  const refreshed = toTokenCache(response);
  await writeJsonSecure(config.tokenCachePath, refreshed);
  return refreshed;
}

export async function getAccessToken(config: AppConfig): Promise<string> {
  const token = await getTokenCache(config);
  if (!token) {
    throw new ConfigError('Not authenticated. Call auth_start_device_login and auth_complete_device_login first.');
  }

  if (token.expiresAt > nowSeconds() + 300) {
    return token.accessToken;
  }

  return (await refreshAccessToken(config, token)).accessToken;
}

export async function exchangeOnBehalfOf(
  config: AppConfig,
  userAssertion: string,
  scopes = config.oboScopes
): Promise<TokenCache> {
  requireOboConfigured(config);
  const key = oboCacheKey(config, userAssertion, scopes);
  const cached = oboCache.get(key);

  if (cached && cached.token.expiresAt > nowSeconds() + 300) {
    return cached.token;
  }

  const response = await getConfidentialClient(config).acquireTokenOnBehalfOf({
    oboAssertion: userAssertion,
    scopes
  });

  if (!response) {
    throw new OAuthProtocolError('MSAL OBO response was empty.');
  }

  const token = toMsalTokenCache(response);
  oboCache.set(key, {
    token,
    assertion: userAssertion,
    scopes: scopes.join(' ')
  });
  return token;
}

export async function getGraphAccessToken(config: AppConfig, scopes = config.oboScopes): Promise<string> {
  const requestContext = getRequestContext();
  if (requestContext?.userAssertion) {
    return (await exchangeOnBehalfOf(config, requestContext.userAssertion, scopes)).accessToken;
  }

  return getAccessToken(config);
}

export async function clearAuth(config: AppConfig): Promise<void> {
  await removeIfExists(config.tokenCachePath);
  await removeIfExists(config.deviceCodeCachePath);
}

export function secondsToIso(seconds?: number): string | undefined {
  return seconds ? new Date(seconds * 1000).toISOString() : undefined;
}
