import { createRemoteJWKSet, jwtVerify, type JWTPayload } from 'jose';
import type { AppConfig } from './config.js';
import { ConfigError } from './config.js';
import { authenticateOAuthBridgeToken } from './oauth-bridge.js';

export type InboundAuthResult = {
  bearerToken?: string;
  userAssertion?: string;
  claims?: JWTPayload;
  disabled: boolean;
};

type OidcMetadata = {
  issuer: string;
  jwks_uri: string;
};

let cachedMetadata: OidcMetadata | undefined;
let cachedJwks: ReturnType<typeof createRemoteJWKSet> | undefined;

export class InboundAuthError extends Error {
  readonly status: number;
  readonly oauthError: string;

  constructor(message: string, status = 401, oauthError = status === 403 ? 'insufficient_scope' : 'invalid_token') {
    super(message);
    this.name = 'InboundAuthError';
    this.status = status;
    this.oauthError = oauthError;
  }
}

function bearerFromHeader(value: string | string[] | undefined): string | undefined {
  const header = Array.isArray(value) ? value[0] : value;
  if (!header) return undefined;
  const match = /^Bearer\s+(.+)$/i.exec(header.trim());
  return match?.[1];
}

async function getOidcMetadata(config: AppConfig): Promise<OidcMetadata> {
  if (cachedMetadata) return cachedMetadata;
  if (!config.oidcMetadataUrl) {
    throw new ConfigError('MS_OIDC_METADATA_URL could not be derived because MS_TENANT_ID is missing.');
  }

  const response = await fetch(config.oidcMetadataUrl, { headers: { Accept: 'application/json' } });
  if (!response.ok) {
    throw new InboundAuthError(`Failed to load OIDC metadata: HTTP ${response.status}`, 500);
  }

  cachedMetadata = (await response.json()) as OidcMetadata;
  return cachedMetadata;
}

async function getRemoteJwks(config: AppConfig) {
  if (cachedJwks) return cachedJwks;
  const metadata = await getOidcMetadata(config);
  cachedJwks = createRemoteJWKSet(new URL(metadata.jwks_uri));
  return cachedJwks;
}

function claimValues(claim: unknown): string[] {
  if (Array.isArray(claim)) {
    return claim.flatMap(value => (typeof value === 'string' ? [value] : []));
  }
  if (typeof claim === 'string') {
    return claim.split(/\s+/).filter(Boolean);
  }
  return [];
}

function finalScopeSegment(scope: string): string {
  const trimmed = scope.replace(/\/+$/, '');
  const slashIndex = trimmed.lastIndexOf('/');
  return slashIndex >= 0 ? trimmed.slice(slashIndex + 1) : trimmed;
}

function hasTokenScope(tokenScopes: Set<string>, requiredScope: string): boolean {
  return tokenScopes.has(requiredScope) || tokenScopes.has(finalScopeSegment(requiredScope));
}

function assertRequiredScopes(config: AppConfig, claims: JWTPayload): void {
  if (config.requiredTokenScopes.length === 0) return;

  const tokenScopes = new Set([
    ...claimValues(claims.scp),
    ...claimValues((claims as { scope?: unknown }).scope),
    ...claimValues((claims as { roles?: unknown }).roles)
  ]);

  const missingScopes = config.requiredTokenScopes.filter(scope => !hasTokenScope(tokenScopes, scope));
  if (missingScopes.length > 0) {
    throw new InboundAuthError(`Access token lacks required scope(s): ${missingScopes.join(', ')}.`, 403);
  }
}

export async function authenticateInboundRequest(
  config: AppConfig,
  authorizationHeader: string | string[] | undefined
): Promise<InboundAuthResult> {
  if (config.inboundAuthDisabled) {
    return { disabled: true };
  }

  const bearerToken = bearerFromHeader(authorizationHeader);
  if (!bearerToken) {
    throw new InboundAuthError('Missing bearer token for remote MCP request.');
  }

  const bridgeAuth = await authenticateOAuthBridgeToken(config, bearerToken);
  if (bridgeAuth) {
    return {
      bearerToken,
      userAssertion: bridgeAuth.userAssertion,
      claims: bridgeAuth.claims,
      disabled: false
    };
  }

  if (config.inboundTokenAudiences.length === 0) {
    throw new ConfigError('MCP_TOKEN_AUDIENCE is required when inbound auth is enabled.');
  }

  const metadata = await getOidcMetadata(config);
  let payload: JWTPayload;
  try {
    const verified = await jwtVerify(bearerToken, await getRemoteJwks(config), {
      audience: config.inboundTokenAudiences,
      issuer: config.inboundTokenIssuer ?? metadata.issuer
    });
    payload = verified.payload;
  } catch (error) {
    const details = error instanceof Error ? error.message : String(error);
    throw new InboundAuthError(`Invalid bearer token: ${details}`);
  }

  assertRequiredScopes(config, payload);

  return {
    bearerToken,
    claims: payload,
    disabled: false
  };
}
