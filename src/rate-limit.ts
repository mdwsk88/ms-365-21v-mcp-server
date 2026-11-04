import type { AppConfig } from './config.js';
import { mountedRoutePaths } from './http-routes.js';

type Window = { count: number; resetAt: number };

const maxTrackedKeys = 10_000;

/** Fixed-window request counter per key (client address or user), kept in memory. */
export class RateLimiter {
  private readonly windows = new Map<string, Window>();

  constructor(
    readonly limit: number,
    private readonly windowMs = 60_000,
    private readonly now: () => number = () => Date.now()
  ) {}

  /** Counts one request. Over the limit, returns how many seconds remain in the window. */
  take(key: string): { allowed: true } | { allowed: false; retryAfterSeconds: number } {
    const now = this.now();
    let window = this.windows.get(key);
    if (!window || window.resetAt <= now) {
      if (!window && this.windows.size >= maxTrackedKeys) this.sweep(now);
      window = { count: 0, resetAt: now + this.windowMs };
      this.windows.set(key, window);
    }
    window.count += 1;
    if (window.count <= this.limit) return { allowed: true };
    return { allowed: false, retryAfterSeconds: Math.max(1, Math.ceil((window.resetAt - now) / 1000)) };
  }

  get size(): number {
    return this.windows.size;
  }

  private sweep(now: number): void {
    for (const [key, window] of this.windows) if (window.resetAt <= now) this.windows.delete(key);
    // Still full: forget the oldest windows rather than grow without bound.
    for (const key of this.windows.keys()) {
      if (this.windows.size < maxTrackedKeys) break;
      this.windows.delete(key);
    }
  }
}

export type RateLimits = {
  /** Endpoints reachable before sign-in, per client address. */
  address: RateLimiter;
  /** Dynamic client registration, per client address: a tenth of the address limit. */
  registration: RateLimiter;
  /** MCP requests per signed-in user. */
  user: RateLimiter;
};

export function createRateLimits(config: AppConfig): RateLimits | undefined {
  if (!config.rateLimitEnabled) return undefined;
  return {
    address: new RateLimiter(config.rateLimitIpPerMinute),
    registration: new RateLimiter(Math.max(1, Math.floor(config.rateLimitIpPerMinute / 10))),
    user: new RateLimiter(config.rateLimitUserPerMinute)
  };
}

/** The client address Express derived from the socket and the trusted proxies (MCP_TRUST_PROXY). */
export function clientAddress(req: any): string {
  return String(req.ip ?? req.socket?.remoteAddress ?? 'unknown');
}

export function rateLimitLog(scope: string, details: Record<string, unknown> = {}): void {
  console.error(JSON.stringify({ component: 'rate_limit', event: 'limited', scope, ...details }));
}

function limitByAddress(limiter: RateLimiter, scope: string) {
  return (req: any, res: any, next: () => void) => {
    const result = limiter.take(clientAddress(req));
    if (result.allowed) {
      next();
      return;
    }
    rateLimitLog(scope, { path: req.path });
    res
      .status(429)
      .set({ 'Retry-After': String(result.retryAfterSeconds), 'Cache-Control': 'no-store' })
      .json({
        error: 'rate_limited',
        error_description: `Too many requests from this address. Retry after ${result.retryAfterSeconds} seconds.`
      });
  };
}

/**
 * Per-address limits for the routes anyone can reach without a session: the OAuth bridge, attachment
 * downloads, confirmation pages and the admin API. Mount before the routes themselves.
 */
export function mountAddressRateLimits(app: any, config: AppConfig, limits: RateLimits | undefined): void {
  if (!limits) return;
  for (const pathname of ['/oauth', '/attachments', '/confirm', '/admin']) {
    for (const routePath of mountedRoutePaths(config, pathname)) app.use(routePath, limitByAddress(limits.address, 'address'));
  }
  for (const routePath of mountedRoutePaths(config, '/oauth/register')) {
    app.use(routePath, limitByAddress(limits.registration, 'registration'));
  }
}
