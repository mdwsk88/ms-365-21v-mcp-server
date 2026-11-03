import type { AppConfig } from './config.js';

/** Bare, lower-case scope name: `https://graph/Mail.Read` → `mail.read`. */
export function normalizedScopeName(scope: string): string {
  const trimmed = scope.trim().replace(/\/+$/, '');
  const slashIndex = trimmed.lastIndexOf('/');
  return (slashIndex >= 0 ? trimmed.slice(slashIndex + 1) : trimmed).toLowerCase();
}

/** Scopes that imply others beyond the plain Read/ReadWrite/.All pattern. */
const impliedScopes: Record<string, string[]> = {
  'sites.fullcontrol.all': ['sites.manage.all', 'sites.readwrite.all', 'sites.read.all'],
  'sites.manage.all': ['sites.readwrite.all', 'sites.read.all'],
  'user.readwrite.all': ['user.read.all', 'user.readbasic.all'],
  'user.read.all': ['user.readbasic.all'],
  'chat.readwrite': ['chat.read', 'chat.readbasic'],
  'chat.read': ['chat.readbasic'],
  'mail.readwrite': ['mail.read', 'mail.readbasic'],
  'mail.read': ['mail.readbasic']
};

/**
 * Whether a granted delegated scope covers a required one: an exact match, `X.ReadWrite[.All|.Shared]`
 * covering `X.Read[...]`, `.All` covering the user-only variant, and the explicit implications above.
 */
export function scopeCovers(granted: string, required: string): boolean {
  const have = normalizedScopeName(granted);
  const need = normalizedScopeName(required);
  if (have === need) return true;
  if (impliedScopes[have]?.includes(need)) return true;
  const parse = (scope: string) => scope.match(/^(.+?)\.(read|readwrite)(?:\.(all|shared))?$/);
  const haveParts = parse(have);
  const needParts = parse(need);
  if (!haveParts || !needParts || haveParts[1] !== needParts[1]) return false;
  const accessCovers = haveParts[2] === 'readwrite' || needParts[2] === 'read';
  const breadthCovers = haveParts[3] === needParts[3] || (haveParts[3] === 'all' && needParts[3] === undefined);
  return accessCovers && breadthCovers;
}

/**
 * Whether this deployment may use a Graph scope. MCP_DISABLED_GRAPH_SCOPES always wins. When
 * MCP_ALLOWED_GRAPH_SCOPES is set, the scope must be covered by one of its entries (fail closed).
 */
export function graphScopeEnabled(
  config: Pick<AppConfig, 'disabledGraphScopes' | 'allowedGraphScopes'>,
  scope: string
): boolean {
  const name = normalizedScopeName(scope);
  if (config.disabledGraphScopes.some(disabled => normalizedScopeName(disabled) === name)) return false;
  if (config.allowedGraphScopes.length === 0) return true;
  return config.allowedGraphScopes.some(allowed => scopeCovers(allowed, name));
}

/**
 * The scope to request from Microsoft Entra for a required scope. With an allow-list, an OBO request
 * for a scope the tenant never consented fails even when a broader one is granted, so ask for the
 * granted covering scope instead.
 */
export function requestableGraphScope(config: Pick<AppConfig, 'allowedGraphScopes'>, scope: string): string {
  if (config.allowedGraphScopes.length === 0) return scope;
  const name = normalizedScopeName(scope);
  if (config.allowedGraphScopes.some(allowed => normalizedScopeName(allowed) === name)) return scope;
  return config.allowedGraphScopes.find(allowed => scopeCovers(allowed, name)) ?? scope;
}

/** Raised when an input needs a Graph scope this deployment has not enabled; reported as scope_not_enabled. */
export class ScopeNotEnabledError extends Error {
  readonly code = 'scope_not_enabled';

  constructor(
    readonly scope: string,
    purpose: string
  ) {
    super(
      `${purpose} needs the Microsoft Graph delegated permission ${scope}, which this deployment has not enabled. ` +
        'An administrator must add and consent it in Microsoft Entra, then allow it in the MCP configuration.'
    );
    this.name = 'ScopeNotEnabledError';
  }
}

export function requireGraphScope(
  config: Pick<AppConfig, 'disabledGraphScopes' | 'allowedGraphScopes'>,
  scope: string,
  purpose: string
): void {
  if (!graphScopeEnabled(config, scope)) throw new ScopeNotEnabledError(scope, purpose);
}
