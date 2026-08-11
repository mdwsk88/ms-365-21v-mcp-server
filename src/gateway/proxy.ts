export type DownstreamAuthMode = 'passthrough' | 'static' | 'none';

export interface DownstreamMcp {
  id: string;
  name: string;
  url: string;
  toolPrefix: string;
  authMode: DownstreamAuthMode;
  staticToken?: string;
  enabled: boolean;
}

/**
 * Parses the reserved downstream configuration format:
 *   id:https://service.example/mcp,other:https://other.example/mcp
 *
 * Proxy forwarding is intentionally not implemented in this iteration. New
 * entries default to user-token passthrough so adding the future transport
 * cannot silently downgrade authentication.
 */
export function parseDownstreamServices(value: string | undefined): DownstreamMcp[] {
  if (!value?.trim()) return [];

  const seen = new Set<string>();
  return value
    .split(',')
    .map((item) => item.trim())
    .filter(Boolean)
    .map((item) => {
      const separator = item.indexOf(':');
      if (separator <= 0) {
        throw new Error('Invalid MCP_DOWNSTREAM_SERVICES entry: expected id:https://host/mcp.');
      }

      const id = item.slice(0, separator).trim();
      const urlValue = item.slice(separator + 1).trim();
      if (!/^[a-z][a-z0-9_-]{1,63}$/i.test(id)) {
        throw new Error('Invalid downstream MCP id: ' + id);
      }
      if (seen.has(id)) {
        throw new Error('Duplicate downstream MCP id: ' + id);
      }

      const url = new URL(urlValue);
      if (url.protocol !== 'https:' && url.protocol !== 'http:') {
        throw new Error('Downstream MCP URL must use HTTP or HTTPS: ' + urlValue);
      }

      seen.add(id);
      return {
        id,
        name: id,
        url: url.toString(),
        toolPrefix: id + '_',
        authMode: 'passthrough' as const,
        enabled: true
      };
    });
}

export function registerDownstreamTools(_services: DownstreamMcp[]): void {
  // Reserved extension point. Proxy discovery and forwarding are a later iteration.
}
