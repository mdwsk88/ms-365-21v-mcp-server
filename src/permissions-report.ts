import type { AppConfig } from './config.js';
import { graphScopeEnabled, normalizedScopeName, scopeCovers } from './scope-policy.js';
import { createToolRegistry } from './tools/index.js';
import type { ToolRegistryEntry } from './tools/types.js';

/**
 * Which Microsoft Graph delegated permissions the configured tool surface needs, what the scope
 * policy hides, and which allowed scopes no tool uses. Built from metadata only: no network access.
 */
export function permissionReport(config: AppConfig) {
  const effective = createToolRegistry(config).getAll().filter(isBusinessTool);
  const unrestricted = createToolRegistry({
    ...config,
    disabledGraphScopes: [],
    allowedGraphScopes: [],
    readOnly: false
  })
    .getAll()
    .filter(isBusinessTool);
  const enabledNames = new Set(effective.map(entry => entry.name));

  const requiredScopes = uniqueScopes(effective.flatMap(entry => entry.graphScopes));
  const requiredNames = new Set(requiredScopes.map(normalizedScopeName));
  const optionalScopes = uniqueScopes(effective.flatMap(entry => entry.optionalGraphScopes ?? [])).filter(
    scope => !requiredNames.has(normalizedScopeName(scope))
  );

  const hiddenTools = unrestricted
    .filter(entry => !enabledNames.has(entry.name))
    .map(entry => ({
      name: entry.name,
      reason: config.readOnly && entry.isWriteOperation ? 'read_only' : 'graph_scope',
      missingScopes:
        config.readOnly && entry.isWriteOperation
          ? []
          : entry.graphScopes.filter(scope => !graphScopeEnabled(config, scope))
    }));

  const usedScopes = [...requiredScopes, ...optionalScopes];
  return {
    readOnly: config.readOnly,
    toolExposureMode: config.toolExposureMode,
    toolCategories: config.toolCategories.length ? config.toolCategories : 'all',
    scopePolicy: {
      allowedGraphScopes: config.allowedGraphScopes.length ? config.allowedGraphScopes : 'all',
      disabledGraphScopes: config.disabledGraphScopes
    },
    enabledToolCount: effective.length,
    requiredScopes,
    optionalScopes: optionalScopes.map(scope => ({
      scope,
      enabled: graphScopeEnabled(config, scope),
      tools: effective.filter(entry => entry.optionalGraphScopes?.includes(scope)).map(entry => entry.name)
    })),
    hiddenTools,
    missingScopesForHiddenTools: uniqueScopes(hiddenTools.flatMap(tool => tool.missingScopes)),
    allowedScopesNotUsed: config.allowedGraphScopes.filter(
      allowed => !usedScopes.some(scope => scopeCovers(allowed, scope))
    ),
    tools: effective.map(entry => ({
      name: entry.name,
      access: entry.isWriteOperation ? 'write' : 'read',
      scopes: entry.graphScopes,
      optionalScopes: entry.optionalGraphScopes?.length ? entry.optionalGraphScopes : undefined
    }))
  };
}

export function formatPermissionReport(report: ReturnType<typeof permissionReport>): string {
  const lines = [
    'Microsoft Graph delegated permissions for this configuration',
    `Exposure: ${report.toolExposureMode}; read-only: ${report.readOnly ? 'yes' : 'no'}; enabled business tools: ${report.enabledToolCount}`,
    '',
    `Required (${report.requiredScopes.length}):`,
    ...report.requiredScopes.map(scope => `  ${scope}`)
  ];
  if (report.optionalScopes.length) {
    lines.push('', 'Optional, needed only for some inputs:');
    for (const item of report.optionalScopes) {
      lines.push(`  ${item.scope} [${item.enabled ? 'enabled' : 'not enabled'}] for ${item.tools.join(', ')}`);
    }
  }
  if (report.hiddenTools.length) {
    lines.push('', `Hidden tools (${report.hiddenTools.length}):`);
    const byReason = new Map<string, string[]>();
    for (const tool of report.hiddenTools) {
      const key = tool.reason === 'read_only' ? 'read-only mode' : `missing ${tool.missingScopes.join(', ')}`;
      byReason.set(key, [...(byReason.get(key) ?? []), tool.name]);
    }
    for (const [reason, names] of byReason) lines.push(`  ${reason}: ${names.join(', ')}`);
  }
  if (report.allowedScopesNotUsed.length) {
    lines.push('', `Allowed but unused by any enabled tool: ${report.allowedScopesNotUsed.join(', ')}`);
  }
  return lines.join('\n');
}

function isBusinessTool(entry: ToolRegistryEntry): boolean {
  return entry.category !== 'gateway';
}

function uniqueScopes(scopes: string[]): string[] {
  const byName = new Map<string, string>();
  for (const scope of scopes) if (!byName.has(normalizedScopeName(scope))) byName.set(normalizedScopeName(scope), scope);
  return [...byName.values()].sort((left, right) => left.localeCompare(right));
}
