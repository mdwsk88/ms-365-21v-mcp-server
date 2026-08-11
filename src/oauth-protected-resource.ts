import type { AppConfig } from './config.js';
import { mountedRoutePaths } from './http-routes.js';

export type ProtectedResourceMetadata = {
  resource: string;
  authorization_servers: string[];
  bearer_methods_supported: string[];
  scopes_supported?: string[];
  resource_documentation?: string;
  resource_name?: string;
};

function quoteHeaderValue(value: string): string {
  return `"${value.replace(/\\/g, '\\\\').replace(/"/g, '\\"')}"`;
}

function unique(values: string[]): string[] {
  return [...new Set(values.filter(Boolean))];
}

export function protectedResourceMetadataPaths(config: AppConfig): string[] {
  const configuredPath = new URL(config.resourceMetadataUrl).pathname;
  return unique([
    configuredPath,
    ...mountedRoutePaths(config, '/.well-known/oauth-protected-resource')
  ]);
}

export function protectedResourceMetadata(config: AppConfig): ProtectedResourceMetadata {
  return {
    resource: config.resourceUrl,
    authorization_servers: config.authorizationServers,
    bearer_methods_supported: ['header'],
    scopes_supported: config.authorizationScopes.length ? config.authorizationScopes : undefined,
    resource_documentation: config.resourceDocumentationUrl,
    resource_name: '21V Microsoft Graph MCP'
  };
}

export function bearerChallenge(
  config: AppConfig,
  options: { error?: string; description?: string; includeScope?: boolean } = {}
): string {
  const params: string[] = [
    `realm=${quoteHeaderValue('mcp')}`,
    `resource_metadata=${quoteHeaderValue(config.resourceMetadataUrl)}`
  ];

  if (options.includeScope !== false && config.authorizationScopes.length > 0) {
    params.push(`scope=${quoteHeaderValue(config.authorizationScopes.join(' '))}`);
  }
  if (options.error) {
    params.push(`error=${quoteHeaderValue(options.error)}`);
  }
  if (options.description) {
    params.push(`error_description=${quoteHeaderValue(options.description)}`);
  }

  return `Bearer ${params.join(', ')}`;
}
