import type { AppConfig, GraphResilienceConfig } from './config.js';
import { getAuditLogger } from './audit-log.js';
import { getRequestContext } from './request-context.js';

type ResilienceEvent = (event: string, details: Record<string, unknown>) => Promise<void> | void;
type Dependencies = {
  sleep?: (milliseconds: number) => Promise<void>;
  now?: () => number;
  onEvent?: ResilienceEvent;
};

export type GraphResiliencePolicy = {
  retryTransientFailures?: boolean;
};

export class GraphCircuitOpenError extends Error {
  readonly code = 'graph_circuit_open';

  constructor(readonly retryAfterMs: number) {
    super(`Microsoft Graph circuit is open. Retry after ${retryAfterMs} ms.`);
    this.name = 'GraphCircuitOpenError';
  }
}

export class GraphRequestTimeoutError extends Error {
  readonly code = 'graph_request_timeout';

  constructor(readonly timeoutMs: number) {
    super(`Microsoft Graph request timed out after ${timeoutMs} ms.`);
    this.name = 'GraphRequestTimeoutError';
  }
}

export class GraphResilience {
  private consecutiveFailures = 0;
  private circuitOpenUntil = 0;
  private readonly sleep: (milliseconds: number) => Promise<void>;
  private readonly now: () => number;
  private readonly onEvent?: ResilienceEvent;

  constructor(
    readonly config: GraphResilienceConfig,
    dependencies: Dependencies = {}
  ) {
    this.sleep = dependencies.sleep ?? ((milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds)));
    this.now = dependencies.now ?? Date.now;
    this.onEvent = dependencies.onEvent;
  }

  async execute(
    operation: (signal: AbortSignal) => Promise<Response>,
    policy: GraphResiliencePolicy = {}
  ): Promise<Response> {
    await this.assertCircuitAllowsRequest();
    const retryTransientFailures = policy.retryTransientFailures ?? true;

    for (let attempt = 0; ; attempt += 1) {
      let response: Response;
      try {
        response = await this.withTimeout(operation);
      } catch (error) {
        if (retryTransientFailures && attempt < this.config.maxRetries) {
          const delayMs = this.backoff(attempt);
          await this.emit('graph_retry', {
            reason: error instanceof GraphRequestTimeoutError ? 'timeout' : 'network_error',
            attempt: attempt + 1,
            delayMs
          });
          await this.sleep(delayMs);
          continue;
        }
        await this.recordFailure(error instanceof Error ? error.name : 'network_error');
        throw error;
      }

      if (response.status === 429 && attempt < this.config.maxRetries) {
        const delayMs =
          retryAfterMilliseconds(response.headers.get('retry-after'), this.now()) ?? this.backoff(attempt);
        await this.emit('graph_retry', { reason: 'http_429', attempt: attempt + 1, delayMs });
        await discardResponse(response);
        await this.sleep(delayMs);
        continue;
      }

      if (
        retryTransientFailures &&
        (response.status === 503 || response.status === 504) &&
        attempt < this.config.maxRetries
      ) {
        const delayMs = this.backoff(attempt);
        await this.emit('graph_retry', { reason: `http_${response.status}`, attempt: attempt + 1, delayMs });
        await discardResponse(response);
        await this.sleep(delayMs);
        continue;
      }

      if (response.status >= 500) {
        await this.recordFailure(`http_${response.status}`);
      } else {
        this.recordSuccess();
      }
      return response;
    }
  }

  snapshot(): object {
    const now = this.now();
    return {
      consecutiveFailures: this.consecutiveFailures,
      circuitOpen: this.circuitOpenUntil > now,
      retryAfterMs: Math.max(0, this.circuitOpenUntil - now)
    };
  }

  private async withTimeout(operation: (signal: AbortSignal) => Promise<Response>): Promise<Response> {
    const controller = new AbortController();
    let timer: NodeJS.Timeout | undefined;
    const timeout = new Promise<never>((_resolve, reject) => {
      timer = setTimeout(() => {
        controller.abort();
        reject(new GraphRequestTimeoutError(this.config.timeoutMs));
      }, this.config.timeoutMs);
      timer.unref?.();
    });
    try {
      return await Promise.race([operation(controller.signal), timeout]);
    } finally {
      if (timer) clearTimeout(timer);
    }
  }

  private backoff(attempt: number): number {
    return Math.max(0, Math.round(this.config.initialBackoffMs * this.config.backoffMultiplier ** attempt));
  }

  private async assertCircuitAllowsRequest(): Promise<void> {
    const now = this.now();
    if (this.circuitOpenUntil > now) {
      const retryAfterMs = this.circuitOpenUntil - now;
      await this.emit('graph_circuit_rejected', { retryAfterMs });
      throw new GraphCircuitOpenError(retryAfterMs);
    }
    if (this.circuitOpenUntil > 0) {
      this.circuitOpenUntil = 0;
      await this.emit('graph_circuit_half_open', {});
    }
  }

  private async recordFailure(reason: string): Promise<void> {
    this.consecutiveFailures += 1;
    if (this.consecutiveFailures >= this.config.circuitBreakerThreshold) {
      this.circuitOpenUntil = this.now() + this.config.circuitBreakerCooldownMs;
      await this.emit('graph_circuit_opened', {
        reason,
        consecutiveFailures: this.consecutiveFailures,
        cooldownMs: this.config.circuitBreakerCooldownMs
      });
    }
  }

  private recordSuccess(): void {
    this.consecutiveFailures = 0;
    this.circuitOpenUntil = 0;
  }

  private async emit(event: string, details: Record<string, unknown>): Promise<void> {
    await this.onEvent?.(event, details);
  }
}

async function discardResponse(response: Response): Promise<void> {
  try {
    await response.body?.cancel();
  } catch {
    // A failed body cleanup must not suppress the configured retry.
  }
}

function retryAfterMilliseconds(value: string | null, now: number): number | undefined {
  if (!value) return undefined;
  const seconds = Number(value);
  if (Number.isFinite(seconds) && seconds >= 0) return Math.min(seconds * 1000, 300_000);
  const date = Date.parse(value);
  if (!Number.isFinite(date)) return undefined;
  return Math.min(Math.max(0, date - now), 300_000);
}

const resilienceByConfig = new WeakMap<AppConfig, GraphResilience>();

async function writeResilienceAudit(config: AppConfig, event: string, details: Record<string, unknown>): Promise<void> {
  const context = getRequestContext();
  const claims = context?.inboundClaims;
  try {
    await getAuditLogger(config).write({
      timestamp: new Date().toISOString(),
      requestId: context?.requestId ?? 'graph-background',
      userId: typeof claims?.oid === 'string' ? claims.oid : (claims?.sub ?? 'local-or-anonymous'),
      userDisplayName: typeof claims?.name === 'string' ? claims.name : undefined,
      toolName: 'graph_resilience',
      toolCategory: 'gateway',
      isWriteOperation: false,
      parameters: details,
      duration: 0,
      success: false,
      errorCode: event,
      eventType: event
    });
  } catch (error) {
    console.error(
      JSON.stringify({
        component: 'graph_resilience',
        event: 'audit_failed',
        error: error instanceof Error ? error.message : String(error)
      })
    );
  }
}

export function getGraphResilience(config: AppConfig): GraphResilience {
  let resilience = resilienceByConfig.get(config);
  if (!resilience) {
    resilience = new GraphResilience(config.graphResilience, {
      onEvent: (event, details) => writeResilienceAudit(config, event, details)
    });
    resilienceByConfig.set(config, resilience);
  }
  return resilience;
}
