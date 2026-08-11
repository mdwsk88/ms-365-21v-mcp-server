import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { AppConfig } from '../config.js';
import { registerDownstreamTools } from '../gateway/proxy.js';
import { authModule } from './auth.js';
import { calendarModule } from './calendar.js';
import { contactsModule } from './contacts.js';
import { confirmModule } from './confirm.js';
import { debugModule } from './debug.js';
import {
  DynamicToolCatalog,
  registerDynamicToolEntries,
  registerDynamicTools
} from './discovery.js';
import { driveModule } from './drive.js';
import { groupsModule } from './groups.js';
import { mailModule } from './mail.js';
import { toolMetadata } from './metadata.js';
import { ToolRegistry, moduleEnabledForDeployment } from './registry.js';
import { instrumentMcpServer, ToolRuntime } from './runtime.js';
import { sharepointModule } from './sharepoint.js';
import { searchModule } from './search.js';
import { smartModule } from './smart.js';
import { teamsModule } from './teams.js';
import type { ToolAccessContext, ToolModule } from './types.js';
import { usersModule } from './users.js';

export const toolModules: ToolModule[] = [
  authModule,
  confirmModule,
  usersModule,
  groupsModule,
  mailModule,
  calendarModule,
  driveModule,
  sharepointModule,
  teamsModule,
  contactsModule,
  searchModule,
  smartModule,
  debugModule
];

export function createToolRegistry(config: AppConfig): ToolRegistry {
  const registry = new ToolRegistry();
  for (const module of toolModules) {
    if (moduleEnabledForDeployment(module, config)) {
      registry.registerModule(module, config, toolMetadata);
    }
  }
  if (config.toolExposureMode !== 'direct') {
    registerDynamicToolEntries(registry);
  }
  return registry;
}

export function createHydratedToolRegistry(config: AppConfig): ToolRegistry {
  const metadataServer = new McpServer({ name: 'ms-365-21v-mcp-registry', version: '0.1.0' });
  return registerTools(metadataServer, config, { bypassRoleFiltering: true });
}

function canRegisterModule(config: AppConfig, module: ToolModule, access: ToolAccessContext): boolean {
  if (!moduleEnabledForDeployment(module, config)) return false;
  if (module.alwaysEnabled || !module.requiredRole) return true;
  if (!config.roleBasedFiltering || access.bypassRoleFiltering) return true;
  const roles = access.userRoles ?? [];
  return roles.includes('mcp.admin') || roles.includes(module.requiredRole);
}

export function registerTools(server: McpServer, config: AppConfig, access: ToolAccessContext = {}): ToolRegistry {
  const registry = createToolRegistry(config);
  const runtime = new ToolRuntime(config, registry, server);
  const dynamicCatalog =
    config.toolExposureMode === 'direct' ? undefined : new DynamicToolCatalog(config);
  instrumentMcpServer(server, runtime, access, dynamicCatalog);

  for (const module of toolModules) {
    const shouldCaptureForDynamicMode =
      dynamicCatalog !== undefined && moduleEnabledForDeployment(module, config);
    if (shouldCaptureForDynamicMode || canRegisterModule(config, module, access)) {
      module.register(server, config);
    }
  }

  if (dynamicCatalog) {
    registerDynamicTools(server, dynamicCatalog, runtime);
  }

  registerDownstreamTools(config.downstreamServices);
  return registry;
}
