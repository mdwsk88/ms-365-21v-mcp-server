import type { JWTPayload } from 'jose';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import { randomUUID } from 'node:crypto';
import { getAuditLogger } from '../audit-log.js';
import type { AppConfig } from '../config.js';
import {
  browserConfirmationResult,
  browserConfirmationStore,
  confirmationUserKey
} from '../browser-confirmation.js';
import {
  createFormElicitor,
  operationRequiresConfirmation,
  requestOperationConfirmation
} from '../confirmation.js';
import { gatewayMetrics } from '../gateway/metrics.js';
import { getRequestContext } from '../request-context.js';
import { errorResult } from './results.js';
import type {
  CapturedToolCallback,
  CapturedToolRegistration,
  ToolAccessContext,
  ToolRegistrationRouter,
  ToolRegistryEntry
} from './types.js';
import { ToolRegistry } from './registry.js';

export function rolesFromClaims(claims: JWTPayload | undefined): string[] {
  const value = claims?.roles;
  if (Array.isArray(value)) {
    return value.filter((role): role is string => typeof role === 'string');
  }
  if (typeof value === 'string') {
    return value.split(/\s+/).filter(Boolean);
  }
  return [];
}

export function hasToolRole(config: AppConfig, entry: ToolRegistryEntry, roles: string[]): boolean {
  if (!config.roleBasedFiltering || config.inboundAuthDisabled || config.transport === 'stdio') return true;
  const requiredRoles = entry.requiredRoles ?? (entry.requiredRole ? [entry.requiredRole] : []);
  if (requiredRoles.length === 0) return true;
  return roles.includes('mcp.admin') || requiredRoles.every(role => roles.includes(role));
}

export class ToolRuntime {
  constructor(
    readonly config: AppConfig,
    readonly registry: ToolRegistry,
    private readonly server: McpServer
  ) {}

  async invoke(
    toolName: string,
    parameters: Record<string, unknown>,
    callback: () => Promise<unknown>,
    extra: unknown = {},
    confirmed = false
  ): Promise<unknown> {
    const entry = this.registry.getByName(toolName);
    if (!entry) {
      return errorResult('tool_not_registered', 'Tool not registered: ' + toolName);
    }

    const startedAt = performance.now();
    const context = getRequestContext();
    const claims = context?.inboundClaims;
    let result: unknown;
    let success = false;
    let errorCode: string | undefined;

    try {
      const roles = rolesFromClaims(claims);
      if (!hasToolRole(this.config, entry, roles)) {
        const requiredRoles = entry.requiredRoles ?? (entry.requiredRole ? [entry.requiredRole] : []);
        result = errorResult(
          'permission_denied',
          `Permission denied: requires all roles [${requiredRoles.join(', ')}].`
        );
      } else if (!confirmed && this.requiresConfirmation(entry)) {
        const confirmation = await requestOperationConfirmation(
          createFormElicitor(this.server, extra as Parameters<typeof createFormElicitor>[1]),
          entry,
          parameters
        );
        if (confirmation.status === 'approved') {
          result = await callback();
        } else if (confirmation.status === 'stopped') {
          result = confirmation.result;
        } else {
          const pending = browserConfirmationStore.create(
            entry,
            parameters,
            confirmationUserKey(claims),
            this.config.confirmTtlSeconds,
            () => this.invoke(toolName, parameters, callback, {}, true) as Promise<CallToolResult>,
            this.config.publicBaseUrl ?? this.config.oauthBridgeIssuer
          );
          result = browserConfirmationResult(pending, confirmation.detail);
        }
      } else {
        result = await callback();
      }
      errorCode = resultErrorCode(result);
      success = !errorCode;
      return result;
    } catch (error) {
      errorCode = errorCodeFromUnknown(error);
      throw error;
    } finally {
      const duration = Math.max(0, Math.round(performance.now() - startedAt));
      gatewayMetrics.record(toolName, entry.category, success, duration);
      try {
        await getAuditLogger(this.config).write({
          timestamp: new Date().toISOString(),
          requestId: context?.requestId ?? randomUUID(),
          userId: claimString(claims?.oid) ?? claims?.sub ?? 'local-or-anonymous',
          userDisplayName:
            claimString(claims?.name) ?? claimString(claims?.preferred_username) ?? claimString(claims?.upn),
          toolName,
          toolCategory: entry.category,
          isWriteOperation: entry.isWriteOperation,
          parameters,
          duration,
          success,
          errorCode,
          graphScopes: entry.graphScopes
        });
      } catch (auditError) {
        console.error(
          JSON.stringify({
            component: 'audit_log',
            event: 'write_failed',
            error: auditError instanceof Error ? auditError.message : String(auditError)
          })
        );
      }
    }
  }

  private requiresConfirmation(entry: ToolRegistryEntry): boolean {
    return Boolean(
      entry.isWriteOperation &&
        entry.requiresConfirmation &&
        entry.operationType &&
        operationRequiresConfirmation(this.config, entry.operationType)
    );
  }
}

function claimString(value: unknown): string | undefined {
  return typeof value === 'string' && value ? value : undefined;
}

function errorCodeFromUnknown(error: unknown): string {
  if (error && typeof error === 'object') {
    const value = (error as { code?: unknown; name?: unknown }).code ?? (error as { name?: unknown }).name;
    if (typeof value === 'string' && value) return value;
  }
  return 'tool_execution_error';
}

function resultErrorCode(result: unknown): string | undefined {
  if (!result || typeof result !== 'object' || !(result as { isError?: unknown }).isError) return undefined;
  const structured = (result as { structuredContent?: unknown }).structuredContent;
  if (structured && typeof structured === 'object') {
    const error = (structured as { error?: unknown }).error;
    if (error && typeof error === 'object') {
      const code = (error as { code?: unknown; name?: unknown }).code ?? (error as { name?: unknown }).name;
      if (typeof code === 'string' && code) return code;
    }
  }
  return 'tool_result_error';
}

export function instrumentMcpServer(
  server: McpServer,
  runtime: ToolRuntime,
  access: ToolAccessContext = {},
  router?: ToolRegistrationRouter
): void {
  const original = server.registerTool.bind(server) as unknown as (
    name: string,
    registration: CapturedToolRegistration,
    callback: CapturedToolCallback
  ) => unknown;

  const wrapped = (
    name: string,
    registration: CapturedToolRegistration,
    callback: CapturedToolCallback
  ) => {
    runtime.registry.updateRegistration(name, registration);
    const entry = runtime.registry.getByName(name);
    if (!entry) {
      return undefined;
    }
    router?.capture(name, entry, registration, callback);
    if (
      !access.bypassRoleFiltering &&
      !hasToolRole(runtime.config, entry, access.userRoles ?? [])
    ) {
      return undefined;
    }
    if (router && !router.shouldExposeDirect(entry)) {
      return undefined;
    }
    if (registration.inputSchema === undefined) {
      const noInputCallback = callback as (extra: unknown) => Promise<unknown> | unknown;
      return original(name, registration, async (extra) =>
        runtime.invoke(name, {}, async () => noInputCallback(extra), extra)
      );
    }
    return original(name, registration, async (args, extra) => {
      const parameters = (args ?? {}) as Record<string, unknown>;
      return runtime.invoke(name, parameters, async () => callback(parameters, extra), extra);
    });
  };

  server.registerTool = wrapped as unknown as typeof server.registerTool;
}
