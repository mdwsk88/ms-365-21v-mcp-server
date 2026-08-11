import { randomUUID } from 'node:crypto';
import { createMcpExpressApp } from '@modelcontextprotocol/sdk/server/express.js';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import { isInitializeRequest } from '@modelcontextprotocol/sdk/types.js';
import type { AppConfig } from './config.js';
import { requireRemoteConfigured } from './config.js';
import { mountAdminRoutes } from './admin-routes.js';
import { mountConfirmationRoutes } from './confirmation-routes.js';
import { authenticateInboundRequest, InboundAuthError, type InboundAuthResult } from './inbound-auth.js';
import { createGraphMcpServer } from './mcp-server.js';
import { mountedRoutePaths } from './http-routes.js';
import { mountOAuthBridgeRoutes } from './oauth-bridge.js';
import {
  bearerChallenge,
  protectedResourceMetadata,
  protectedResourceMetadataPaths
} from './oauth-protected-resource.js';
import { runWithRequestContext } from './request-context.js';
import { rolesFromClaims } from './tools/runtime.js';

function jsonError(status: number, message: string) {
  return {
    status,
    body: {
      jsonrpc: '2.0',
      error: {
        code: status === 401 ? -32001 : -32603,
        message
      },
      id: null
    }
  };
}

function oauthError(status: number, error: string, description: string) {
  return {
    status,
    body: {
      error,
      error_description: description
    }
  };
}

function headerValue(value: string | string[] | undefined): string | undefined {
  return Array.isArray(value) ? value[0] : value;
}

function hasBearerAuthorization(value: string | string[] | undefined): boolean {
  return /^Bearer\s+.+$/i.test(headerValue(value)?.trim() ?? '');
}

const unauthenticatedDiscoveryMethods = new Set([
  'initialize',
  'notifications/initialized',
  'ping',
  'tools/list',
  'prompts/list',
  'resources/list',
  'resources/templates/list'
]);

function isUnauthenticatedDiscoveryRequest(body: unknown): boolean {
  const messages = Array.isArray(body) ? body : [body];
  if (messages.length === 0) return false;

  return messages.every(message => {
    if (!message || typeof message !== 'object') return false;
    const method = (message as { method?: unknown }).method;
    return typeof method === 'string' && unauthenticatedDiscoveryMethods.has(method);
  });
}

function shouldAllowUnauthenticatedDiscovery(
  config: AppConfig,
  req: any,
  routeAllowsUnauthenticatedDiscovery: boolean
): boolean {
  return (
    routeAllowsUnauthenticatedDiscovery &&
    config.allowUnauthenticatedDiscovery &&
    !hasBearerAuthorization(req.headers.authorization) &&
    isUnauthenticatedDiscoveryRequest(req.body)
  );
}

function isInitializationRequest(body: unknown): boolean {
  const messages = Array.isArray(body) ? body : [body];
  return messages.some(message => isInitializeRequest(message));
}

function attachAuth(req: any, auth: InboundAuthResult): void {
  if (auth.disabled) return;

  req.auth = {
    token: auth.bearerToken ?? '',
    clientId: String(auth.claims?.azp ?? auth.claims?.appid ?? auth.claims?.aud ?? ''),
    scopes: String(auth.claims?.scp ?? '').split(/\s+/).filter(Boolean),
    expiresAt: typeof auth.claims?.exp === 'number' ? auth.claims.exp : undefined,
    extra: { claims: auth.claims }
  };
}

type McpHttpSession = {
  server: ReturnType<typeof createGraphMcpServer>;
  transport: StreamableHTTPServerTransport;
};

export async function startHttpServer(config: AppConfig): Promise<void> {
  if (!config.inboundAuthDisabled) {
    requireRemoteConfigured(config);
  }

  const app = createMcpExpressApp({ host: config.httpHost });
  const sessions = new Map<string, McpHttpSession>();
  const closingTransports = new WeakSet<StreamableHTTPServerTransport>();

  mountAdminRoutes(app, config);

  for (const healthPath of mountedRoutePaths(config, '/healthz')) {
    app.get(healthPath, (_req: any, res: any) => {
      res.json({
        ok: true,
        name: 'ms-365-21v-mcp-server',
        mode: 'http',
        auth: config.inboundAuthDisabled ? 'disabled' : 'bearer-jwt-obo',
        graphBaseUrl: config.graphBaseUrl,
        resourceUrl: config.resourceUrl,
        resourceMetadataUrl: config.resourceMetadataUrl,
        catalogHttpPath: config.httpPath,
        oauthHttpPath: config.oauthHttpPath ?? config.httpPath,
        microsoftLoginClientType: config.oauthBridgeEnabled
          ? config.oauthBridgeMicrosoftClientType
          : undefined,
        toolExposureMode: config.toolExposureMode,
        sendMode: config.sendMode,
        allowUnauthenticatedDiscovery: config.allowUnauthenticatedDiscovery,
        stateless: config.httpStateless,
        disabledGraphScopes: config.disabledGraphScopes
      });
    });
  }

  for (const metadataPath of protectedResourceMetadataPaths(config)) {
    app.get(metadataPath, (_req: any, res: any) => {
      res.json(protectedResourceMetadata(config));
    });
  }

  mountOAuthBridgeRoutes(app, config);
  mountConfirmationRoutes(app, config);

  const handleMcpRequest = async (
    req: any,
    res: any,
    routeAllowsUnauthenticatedDiscovery: boolean
  ) => {
    if (!['POST', 'GET', 'DELETE'].includes(req.method)) {
      res.status(405).set('Allow', 'GET, POST, DELETE').json(jsonError(405, 'Method not allowed.').body);
      return;
    }

    try {
      if (req.method === 'GET' || req.method === 'DELETE') {
        const auth = await authenticateInboundRequest(config, req.headers.authorization);
        const sessionId = headerValue(req.headers['mcp-session-id']);
        const existingSession = sessionId ? sessions.get(sessionId) : undefined;
        if (!existingSession) {
          res.status(404).json(jsonError(404, 'MCP session not found.').body);
          return;
        }
        attachAuth(req, auth);
        await runWithRequestContext(
          {
            requestId: randomUUID(),
            userAssertion: auth.userAssertion ?? auth.bearerToken,
            inboundClaims: auth.claims
          },
          async () => existingSession.transport.handleRequest(req, res)
        );
        return;
      }

      const anonymousDiscovery = shouldAllowUnauthenticatedDiscovery(
        config,
        req,
        routeAllowsUnauthenticatedDiscovery
      );
      const auth = anonymousDiscovery
        ? { disabled: true }
        : await authenticateInboundRequest(config, req.headers.authorization);
      const sessionId = config.httpStateless ? undefined : headerValue(req.headers['mcp-session-id']);
      const existingSession = sessionId ? sessions.get(sessionId) : undefined;
      const isInit = !config.httpStateless && isInitializationRequest(req.body);
      const server =
        existingSession?.server ??
        createGraphMcpServer(config, {
          userRoles: rolesFromClaims(auth.claims),
          bypassRoleFiltering: config.inboundAuthDisabled || anonymousDiscovery
        });
      const transport =
        existingSession?.transport ??
        new StreamableHTTPServerTransport({
          sessionIdGenerator: !config.httpStateless && isInit ? () => randomUUID() : undefined,
          enableJsonResponse: false,
          onsessioninitialized: config.httpStateless
            ? undefined
            : initializedSessionId => {
                sessions.set(initializedSessionId, { server, transport });
              }
        });
      transport.onclose = () => {
        if (closingTransports.has(transport)) return;
        closingTransports.add(transport);
        const initializedSessionId = transport.sessionId;
        if (initializedSessionId) sessions.delete(initializedSessionId);
        void server.close();
      };

      attachAuth(req, auth);

      await runWithRequestContext(
        {
          requestId: randomUUID(),
          userAssertion: auth.userAssertion ?? auth.bearerToken,
          inboundClaims: auth.claims
        },
        async () => {
          if (!existingSession) {
            await server.connect(transport);
          }
          await transport.handleRequest(req, res, req.body);
        }
      );

      if (!existingSession && (config.httpStateless || !isInit)) {
        res.on('close', () => {
          void transport.close();
        });
      }
    } catch (error) {
      const response =
        error instanceof InboundAuthError
          ? oauthError(error.status, error.oauthError, error.message)
          : jsonError(500, error instanceof Error ? error.message : 'Internal server error');

      if (!res.headersSent) {
        if (error instanceof InboundAuthError) {
          res.set(
            'WWW-Authenticate',
            bearerChallenge(config, {
              error: error.oauthError,
              description: error.message
            })
          );
        }
        res.status(response.status).json(response.body);
      }
    }
  };

  const mcpRoutes = new Map<string, boolean>();
  for (const mcpPath of mountedRoutePaths(config, config.httpPath)) {
    mcpRoutes.set(mcpPath, true);
  }
  if (config.oauthHttpPath) {
    for (const mcpPath of mountedRoutePaths(config, config.oauthHttpPath)) {
      mcpRoutes.set(mcpPath, false);
    }
  }
  for (const [mcpPath, routeAllowsUnauthenticatedDiscovery] of mcpRoutes) {
    app.all(mcpPath, (req: any, res: any) =>
      handleMcpRequest(req, res, routeAllowsUnauthenticatedDiscovery)
    );
  }

  function hostForLog(host: string): string {
    return host.includes(':') ? `[${host}]` : host;
  }

  function listen(host: string, required: boolean): void {
    const server = app.listen(config.httpPort, host, () => {
      console.error(
        `MS 365-21V MCP Server listening on http://${hostForLog(host)}:${config.httpPort}${config.httpPath}`
      );
      if (config.oauthHttpPath) {
        console.error(`OAuth-required MCP endpoint: ${config.resourceUrl}`);
      }
      console.error(`OAuth protected resource metadata: ${config.resourceMetadataUrl}`);
    });
    server.on('error', (error: Error) => {
      console.error(`Failed to listen on ${host}:${config.httpPort}:`, error);
      if (required) process.exit(1);
    });
  }

  listen(config.httpHost, true);

  if (config.httpHost === '127.0.0.1') {
    listen('::1', false);
  }
}
