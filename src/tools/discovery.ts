import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import { z } from 'zod/v4';
import type { AppConfig } from '../config.js';
import { getRequestContext } from '../request-context.js';
import { describeTool, errorResult, jsonResult } from './results.js';
import { hasToolRole, rolesFromClaims, type ToolRuntime } from './runtime.js';
import type {
  CapturedToolCallback,
  CapturedToolRegistration,
  ToolRegistrationRouter,
  ToolRegistryEntry
} from './types.js';
import type { ToolRegistry } from './registry.js';

export const dynamicToolNames = [
  'gateway_search_tools',
  'gateway_get_tool_schema',
  'gateway_execute_tool'
] as const;

type CapturedTool = {
  entry: ToolRegistryEntry;
  registration: CapturedToolRegistration;
  callback: CapturedToolCallback;
  schema: z.ZodType;
};

function schemaFromRegistration(registration: CapturedToolRegistration): z.ZodType {
  const input = registration.inputSchema;
  if (!input) return z.object({});
  if (typeof input === 'object' && input !== null && '_zod' in input) {
    return input as z.ZodType;
  }
  return z.object(input as Record<string, z.ZodType>);
}

function normalize(value: string): string {
  return value.normalize('NFKC').toLowerCase().replace(/[_-]+/g, ' ').replace(/\s+/g, ' ').trim();
}

function tokens(value: string): string[] {
  const normalized = normalize(value);
  const result = new Set(normalized.match(/[a-z0-9]+|[\p{Script=Han}]+/gu) ?? []);
  for (const segment of [...result]) {
    if (!/^[\p{Script=Han}]+$/u.test(segment)) continue;
    const characters = Array.from(segment);
    for (let index = 0; index < characters.length - 1; index += 1) {
      result.add(`${characters[index]}${characters[index + 1]}`);
    }
  }
  return [...result];
}

function scoreTool(query: string, tool: CapturedTool): number {
  const normalizedQuery = normalize(query);
  if (!normalizedQuery) return 1;
  const name = normalize(tool.entry.name);
  const title = normalize(tool.entry.title ?? '');
  const description = normalize(tool.entry.description);
  const category = normalize(tool.entry.category);
  let score = 0;
  if (name === normalizedQuery) score += 500;
  else if (name.includes(normalizedQuery)) score += 240;
  if (title === normalizedQuery) score += 300;
  else if (title.includes(normalizedQuery)) score += 160;
  if (description.includes(normalizedQuery)) score += 120;
  if (category === normalizedQuery) score += 80;
  for (const token of tokens(normalizedQuery)) {
    if (name.split(' ').includes(token)) score += 45;
    else if (name.includes(token)) score += 25;
    if (title.includes(token)) score += 20;
    if (description.includes(token)) score += 10;
    if (category.includes(token)) score += 8;
  }
  return score;
}

function validationMessage(error: z.ZodError): string {
  return error.issues
    .slice(0, 10)
    .map(issue => `${issue.path.join('.') || '<root>'}: ${issue.message}`)
    .join('; ');
}

export function registerDynamicToolEntries(registry: ToolRegistry): void {
  for (const name of dynamicToolNames) {
    registry.register({
      name,
      category: 'gateway',
      module: 'Dynamic Tool Discovery',
      description: '',
      isWriteOperation: false,
      requiresConfirmation: false,
      graphScopes: [],
      requiredRoles: []
    });
  }
}

export class DynamicToolCatalog implements ToolRegistrationRouter {
  private readonly tools = new Map<string, CapturedTool>();

  constructor(readonly config: AppConfig) {}

  capture(
    name: string,
    entry: ToolRegistryEntry,
    registration: CapturedToolRegistration,
    callback: CapturedToolCallback
  ): void {
    if (entry.category === 'gateway') return;
    this.tools.set(name, {
      entry: { ...entry },
      registration: { ...registration },
      callback,
      schema: schemaFromRegistration(registration)
    });
  }

  shouldExposeDirect(entry: ToolRegistryEntry): boolean {
    if (entry.category === 'gateway') return true;
    if (this.config.toolExposureMode === 'discovery') return false;
    if (this.config.toolExposureMode === 'direct') return true;
    return (
      this.config.directTools.includes(entry.name.toLowerCase()) ||
      this.config.directToolCategories.includes(entry.category.toLowerCase())
    );
  }

  search(query = '', category?: string, limit = 8): object {
    const normalizedCategory = normalize(category ?? '');
    const ranked = this.visibleTools()
      .filter(tool => !normalizedCategory || normalize(tool.entry.category) === normalizedCategory)
      .map(tool => ({ tool, score: scoreTool(query, tool) }))
      .filter(result => !query.trim() || result.score > 0)
      .sort((left, right) => right.score - left.score || left.tool.entry.name.localeCompare(right.tool.entry.name));
    const selected = ranked.slice(0, Math.min(Math.max(limit, 1), 20));
    return {
      query,
      category: category || undefined,
      found: selected.length,
      totalAvailable: ranked.length,
      tools: selected.map(({ tool, score }) => ({
        name: tool.entry.name,
        title: tool.entry.title,
        category: tool.entry.category,
        description: tool.entry.description,
        readOnly: !tool.entry.isWriteOperation,
        direct: this.shouldExposeDirect(tool.entry),
        score
      })),
      next: 'Call gateway_get_tool_schema with the exact tool name, then gateway_execute_tool.'
    };
  }

  describe(name: string): object | undefined {
    const tool = this.visibleTool(name);
    if (!tool) return undefined;
    return {
      name: tool.entry.name,
      title: tool.entry.title,
      category: tool.entry.category,
      description: tool.entry.description,
      inputSchema: z.toJSONSchema(tool.schema, { target: 'draft-7' }),
      annotations: tool.registration.annotations,
      readOnly: !tool.entry.isWriteOperation,
      requiresConfirmation: tool.entry.requiresConfirmation,
      operationType: tool.entry.operationType,
      graphScopes: tool.entry.graphScopes,
      requiredRoles: tool.entry.requiredRoles,
      direct: this.shouldExposeDirect(tool.entry)
    };
  }

  async execute(
    name: string,
    parameters: Record<string, unknown>,
    extra: unknown,
    runtime: ToolRuntime
  ): Promise<unknown> {
    const tool = this.visibleTool(name);
    if (!tool) {
      return errorResult('tool_not_available', 'Tool is unknown or not authorized for the signed-in user.');
    }
    const parsed = await tool.schema.safeParseAsync(parameters);
    if (!parsed.success) {
      return errorResult('invalid_tool_arguments', validationMessage(parsed.error));
    }
    const callback = async () =>
      tool.registration.inputSchema === undefined
        ? tool.callback(extra)
        : tool.callback(parsed.data as Record<string, unknown>, extra);
    return runtime.invoke(name, parsed.data as Record<string, unknown>, callback, extra);
  }

  private visibleTools(): CapturedTool[] {
    const roles = rolesFromClaims(getRequestContext()?.inboundClaims);
    return [...this.tools.values()].filter(tool => hasToolRole(this.config, tool.entry, roles));
  }

  private visibleTool(name: string): CapturedTool | undefined {
    const tool = this.tools.get(name);
    if (!tool) return undefined;
    const roles = rolesFromClaims(getRequestContext()?.inboundClaims);
    return hasToolRole(this.config, tool.entry, roles) ? tool : undefined;
  }
}

export function registerDynamicTools(
  server: McpServer,
  catalog: DynamicToolCatalog,
  runtime: ToolRuntime
): void {
  server.registerTool(
    'gateway_search_tools',
    {
      title: 'Search Available Tools',
      description: describeTool(
        'Search the signed-in user\'s authorized Microsoft 365 tools by English or Chinese intent. Use this when no directly exposed tool clearly matches the request.',
        ['搜索可用工具', '查找能完成任务的工具', '发现更多Microsoft 365能力']
      ),
      inputSchema: {
        query: z.string().max(300).default('').describe('English or Chinese task description.'),
        category: z.string().max(50).optional().describe('Optional exact category filter, such as mail or drive.'),
        limit: z.number().int().min(1).max(20).default(8).describe('Maximum matching tools to return.')
      },
      annotations: { readOnlyHint: true, openWorldHint: false }
    },
    async ({ query, category, limit }) => jsonResult(catalog.search(query, category, limit))
  );

  server.registerTool(
    'gateway_get_tool_schema',
    {
      title: 'Get Tool Schema',
      description: describeTool(
        'Get the exact input schema and authorization metadata for a tool returned by gateway_search_tools.',
        ['查看工具参数', '获取工具Schema', '确认工具需要哪些参数']
      ),
      inputSchema: {
        toolName: z.string().min(1).max(150).describe('Exact tool name returned by gateway_search_tools.')
      },
      annotations: { readOnlyHint: true, openWorldHint: false }
    },
    async ({ toolName }) => {
      const schema = catalog.describe(toolName);
      return schema
        ? jsonResult(schema)
        : errorResult('tool_not_available', 'Tool is unknown or not authorized for the signed-in user.');
    }
  );

  server.registerTool(
    'gateway_execute_tool',
    {
      title: 'Execute Authorized Tool',
      description: describeTool(
        'Execute a tool returned by gateway_search_tools using parameters from gateway_get_tool_schema. The target tool is rechecked for App Roles, confirmation policy, audit, and Graph scope controls.',
        ['执行搜索到的工具', '调用长尾工具', '执行Microsoft 365操作']
      ),
      inputSchema: {
        toolName: z.string().min(1).max(150).describe('Exact authorized tool name.'),
        parameters: z.record(z.string(), z.unknown()).default({}).describe('Arguments matching gateway_get_tool_schema.')
      },
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true }
    },
    async ({ toolName, parameters }, extra) =>
      (await catalog.execute(toolName, parameters, extra, runtime)) as CallToolResult
  );
}
