import type { AppConfig } from '../config.js';
import type { ToolMetadata, ToolModule, ToolRegistryEntry } from './types.js';

export class ToolRegistry {
  private readonly entries = new Map<string, ToolRegistryEntry>();

  register(entry: ToolRegistryEntry): void {
    if (this.entries.has(entry.name)) {
      throw new Error('Duplicate MCP tool registration: ' + entry.name);
    }
    const requiredRoles = uniqueRoles(entry.requiredRoles ?? (entry.requiredRole ? [entry.requiredRole] : []));
    this.entries.set(entry.name, cloneEntry({
      ...entry,
      requiredRole: entry.requiredRole ?? requiredRoles[0],
      requiredRoles
    }));
  }

  registerModule(module: ToolModule, config: AppConfig, metadataOverrides: Record<string, ToolMetadata> = {}): void {
    const toolNames = module.activeToolNames?.(config) ?? module.toolNames;
    for (const name of toolNames) {
      const metadata: ToolMetadata = {
        ...(module.toolMetadata?.[name] ?? {}),
        ...(metadataOverrides[name] ?? {})
      };
      if (toolDisabledByGraphScopes(metadata.graphScopes ?? [], config.disabledGraphScopes)) {
        continue;
      }
      const requiredRoles = uniqueRoles([
        ...(module.requiredRole ? [module.requiredRole] : []),
        ...(metadata.requiredRoles ?? [])
      ]);
      this.register({
        name,
        category: module.category,
        module: module.displayName,
        description: '',
        isWriteOperation: metadata.isWriteOperation ?? false,
        requiresConfirmation: metadata.requiresConfirmation ?? metadata.isWriteOperation ?? false,
        operationType: metadata.operationType,
        graphScopes: metadata.graphScopes ?? [],
        requiredRole: requiredRoles[0],
        requiredRoles
      });
    }
  }

  updateRegistration(name: string, registration: { title?: string; description?: string }): void {
    const entry = this.entries.get(name);
    if (!entry) return;
    entry.title = registration.title ?? entry.title;
    entry.description = registration.description ?? entry.description;
  }

  getByName(name: string): ToolRegistryEntry | undefined {
    const entry = this.entries.get(name);
    return entry ? cloneEntry(entry) : undefined;
  }

  getByCategory(category: string): ToolRegistryEntry[] {
    return [...this.entries.values()]
      .filter((entry) => entry.category === category)
      .map(cloneEntry);
  }

  getAllCategories(): string[] {
    return [...new Set([...this.entries.values()].map((entry) => entry.category))].sort();
  }

  getWriteOperations(): ToolRegistryEntry[] {
    return [...this.entries.values()]
      .filter((entry) => entry.isWriteOperation)
      .map(cloneEntry);
  }

  getAll(): ToolRegistryEntry[] {
    return [...this.entries.values()].map(cloneEntry);
  }

  toJSON(): object {
    return {
      categories: this.getAllCategories(),
      tools: this.getAll().sort((left, right) => left.name.localeCompare(right.name))
    };
  }
}

export function toolDisabledByGraphScopes(graphScopes: string[], disabledGraphScopes: string[]): boolean {
  if (graphScopes.length === 0 || disabledGraphScopes.length === 0) return false;
  const disabled = new Set(disabledGraphScopes.map(normalizedScopeName));
  return graphScopes.some(scope => disabled.has(normalizedScopeName(scope)));
}

function normalizedScopeName(scope: string): string {
  const trimmed = scope.trim().replace(/\/+$/, '');
  const slashIndex = trimmed.lastIndexOf('/');
  return (slashIndex >= 0 ? trimmed.slice(slashIndex + 1) : trimmed).toLowerCase();
}

function uniqueRoles(roles: string[]): string[] {
  return [...new Set(roles.map(role => role.trim()).filter(Boolean))];
}

function cloneEntry(entry: ToolRegistryEntry): ToolRegistryEntry {
  return {
    ...entry,
    graphScopes: [...entry.graphScopes],
    requiredRoles: [...(entry.requiredRoles ?? [])]
  };
}

export function moduleEnabledForDeployment(module: ToolModule, config: AppConfig): boolean {
  if (module.isEnabled && !module.isEnabled(config)) return false;
  if (module.alwaysEnabled) return true;
  return config.toolCategories.length === 0 || config.toolCategories.includes(module.category);
}
