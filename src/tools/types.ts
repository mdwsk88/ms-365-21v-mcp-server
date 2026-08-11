import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { AppConfig } from '../config.js';

export interface ToolRegistryEntry {
  name: string;
  category: string;
  module: string;
  title?: string;
  description: string;
  isWriteOperation: boolean;
  requiresConfirmation: boolean;
  operationType?: string;
  graphScopes: string[];
  requiredRole?: string;
  requiredRoles?: string[];
}

export type ToolMetadata = Partial<
  Pick<
    ToolRegistryEntry,
    'isWriteOperation' | 'requiresConfirmation' | 'operationType' | 'graphScopes' | 'requiredRoles'
  >
>;

export interface ToolModule {
  category: string;
  displayName: string;
  description: string;
  requiredRole?: string;
  toolNames: string[];
  activeToolNames?(config: AppConfig): string[];
  toolMetadata?: Record<string, ToolMetadata>;
  alwaysEnabled?: boolean;
  isEnabled?(config: AppConfig): boolean;
  register(server: McpServer, config: AppConfig): void;
}

export interface ToolAccessContext {
  userRoles?: string[];
  bypassRoleFiltering?: boolean;
}

export type CapturedToolRegistration = {
  title?: string;
  description?: string;
  inputSchema?: unknown;
  outputSchema?: unknown;
  annotations?: unknown;
  _meta?: Record<string, unknown>;
};

export type CapturedToolCallback = (...args: unknown[]) => Promise<unknown> | unknown;

export interface ToolRegistrationRouter {
  capture(
    name: string,
    entry: ToolRegistryEntry,
    registration: CapturedToolRegistration,
    callback: CapturedToolCallback
  ): void;
  shouldExposeDirect(entry: ToolRegistryEntry): boolean;
}
