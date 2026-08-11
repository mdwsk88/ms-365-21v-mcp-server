import { timingSafeEqual } from 'node:crypto';
import type { AppConfig } from './config.js';
import { getAuditLogger } from './audit-log.js';
import { gatewayMetrics } from './gateway/metrics.js';
import { getGraphResilience } from './graph-resilience.js';
import { mountedRoutePaths } from './http-routes.js';
import { createHydratedToolRegistry } from './tools/index.js';

function bearerToken(value: string | string[] | undefined): string | undefined {
  const header = Array.isArray(value) ? value[0] : value;
  const match = /^Bearer\s+(.+)$/i.exec(header?.trim() ?? '');
  return match?.[1];
}

function tokensEqual(actual: string | undefined, expected: string): boolean {
  if (!actual) return false;
  const actualBuffer = Buffer.from(actual);
  const expectedBuffer = Buffer.from(expected);
  return actualBuffer.length === expectedBuffer.length && timingSafeEqual(actualBuffer, expectedBuffer);
}

export function mountAdminRoutes(app: any, config: AppConfig): void {
  const registry = createHydratedToolRegistry(config);

  const authenticate = (req: any, res: any, next: () => void) => {
    if (!config.adminToken) {
      res.status(503).json({ error: 'admin_api_not_configured' });
      return;
    }
    if (!tokensEqual(bearerToken(req.headers.authorization), config.adminToken)) {
      res.set('WWW-Authenticate', 'Bearer realm="mcp-admin"');
      res.status(401).json({ error: 'invalid_admin_token' });
      return;
    }
    next();
  };

  const mountGet = (pathname: string, handler: (req: any, res: any) => void | Promise<void>) => {
    for (const routePath of mountedRoutePaths(config, pathname)) {
      app.get(routePath, authenticate, handler);
    }
  };

  mountGet('/admin/health', (_req, res) => {
    res.json({
      ok: true,
      name: 'ms-365-21v-mcp-server',
      transport: config.transport,
      toolCount: registry.getAll().length,
      categories: registry.getAllCategories(),
      roleBasedFiltering: config.roleBasedFiltering,
      auditLogEnabled: config.auditLogEnabled,
      downstreamServiceCount: config.downstreamServices.filter((service) => service.enabled).length,
      graphResilience: getGraphResilience(config).snapshot(),
      metrics: gatewayMetrics.snapshot()
    });
  });

  mountGet('/admin/tools', (_req, res) => {
    res.json(registry.toJSON());
  });

  mountGet('/admin/tools/categories', (_req, res) => {
    res.json({
      categories: registry.getAllCategories().map((category) => ({
        category,
        toolCount: registry.getByCategory(category).length
      }))
    });
  });

  mountGet('/admin/audit-summary', async (_req, res) => {
    res.json(await getAuditLogger(config).summary(7));
  });

  mountGet('/admin/metrics', (_req, res) => {
    res.type('text/plain; version=0.0.4; charset=utf-8').send(gatewayMetrics.prometheus());
  });
}
