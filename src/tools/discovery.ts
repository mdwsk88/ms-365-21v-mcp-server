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
  return [...result].filter(token => !stopwords.has(token));
}

const stopwords = new Set([
  'a', 'an', 'and', 'are', 'as', 'at', 'be', 'by', 'chinese', 'delegated', 'for', 'from', 'in', 'intents',
  'is', 'it', 'me', 'my', 'of', 'on', 'or', 'requires', 'the', 'this', 'to', 'use', 'with', '中文意图'
]);

/** Bilingual intent groups; a query word also matches the other words of its group. */
const synonymGroups: string[][] = [
  ['mail', 'email', 'emails', '邮件', '邮箱'],
  ['message', 'messages', '消息'],
  ['chat', 'chats', '聊天', '群聊', '私聊', '会话'],
  ['team', '团队'],
  ['channel', 'channels', '频道'],
  ['calendar', 'event', 'events', 'meeting', 'meetings', '日历', '日程', '会议'],
  ['file', 'files', 'document', 'documents', 'doc', 'docs', '文件', '文档', '资料'],
  ['drive', 'onedrive', '网盘'],
  ['sharepoint', 'site', 'sites', '站点', '网站'],
  ['folder', 'folders', '文件夹', '目录'],
  ['attachment', 'attachments', '附件'],
  ['contact', 'contacts', '联系人', '通讯录'],
  ['user', 'users', 'person', 'people', 'colleague', '用户', '同事', '员工'],
  ['group', 'groups', '群组'],
  ['send', '发送'],
  ['reply', '回复'],
  ['forward', '转发'],
  ['delete', 'remove', '删除', '移除', '撤回'],
  ['create', 'new', 'add', '创建', '新建', '添加'],
  ['update', 'modify', 'rename', 'edit', '修改', '更新', '重命名'],
  ['search', 'find', '搜索', '查找'],
  ['download', '下载'],
  ['upload', '上传'],
  ['share', 'sharing', 'link', '分享', '共享', '链接'],
  ['permission', 'permissions', '权限'],
  ['version', 'versions', '版本'],
  ['member', 'members', '成员'],
  ['flag', 'flagged', '标旗', '旗标', '跟进'],
  ['unread', '未读'],
  ['schedule', 'free', 'busy', 'availability', '空闲', '忙闲', '有空'],
  ['draft', 'drafts', '草稿'],
  ['list', 'lists', '列表'],
  ['item', 'items', '条目', '记录'],
  ['read', '已读', '阅读'],
  ['move', '移动'],
  ['copy', '复制'],
  ['manager', '经理', '上级', '领导'],
  ['respond', 'accept', 'decline', '接受', '拒绝', '答复'],
  ['cancel', '取消']
];

const synonymIndex = new Map<string, string[]>();
for (const group of synonymGroups) {
  for (const word of group) {
    synonymIndex.set(word, [...new Set([...(synonymIndex.get(word) ?? []), ...group.filter(other => other !== word)])]);
  }
}

/** Query terms with weights: original words 1, synonyms 0.6. */
export function expandQuery(query: string): Map<string, number> {
  const terms = new Map<string, number>();
  for (const token of tokens(query)) terms.set(token, 1);
  for (const token of [...terms.keys()]) {
    for (const synonym of synonymIndex.get(token) ?? []) {
      for (const synonymToken of tokens(synonym)) {
        if (!terms.has(synonymToken)) terms.set(synonymToken, 0.6);
      }
    }
  }
  return terms;
}

type IndexedTool = { tool: CapturedTool; frequencies: Map<string, number>; length: number };
type ToolIndex = { tools: Map<string, IndexedTool>; documentFrequency: Map<string, number>; averageLength: number };

const fieldWeights = { name: 3, title: 2, category: 2, description: 1 } as const;

function indexTools(tools: CapturedTool[]): ToolIndex {
  const indexed = new Map<string, IndexedTool>();
  const documentFrequency = new Map<string, number>();
  let totalLength = 0;
  for (const tool of tools) {
    const frequencies = new Map<string, number>();
    const fields: Array<[string, number]> = [
      [tool.entry.name, fieldWeights.name],
      [tool.entry.title ?? '', fieldWeights.title],
      [tool.entry.category, fieldWeights.category],
      [tool.entry.description, fieldWeights.description]
    ];
    let length = 0;
    for (const [value, weight] of fields) {
      for (const token of tokens(value)) {
        frequencies.set(token, (frequencies.get(token) ?? 0) + weight);
        length += weight;
      }
    }
    for (const token of frequencies.keys()) documentFrequency.set(token, (documentFrequency.get(token) ?? 0) + 1);
    totalLength += length;
    indexed.set(tool.entry.name, { tool, frequencies, length });
  }
  return { tools: indexed, documentFrequency, averageLength: tools.length ? totalLength / tools.length : 1 };
}

/** BM25 over weighted name/title/category/description fields, plus exact-name and phrase boosts. */
function scoreTool(query: string, terms: Map<string, number>, indexed: IndexedTool, index: ToolIndex): number {
  const normalizedQuery = normalize(query);
  if (!normalizedQuery) return 1;
  const k1 = 1.2;
  const b = 0.75;
  const documents = index.tools.size;
  let score = 0;
  for (const [term, weight] of terms) {
    const frequency = indexed.frequencies.get(term);
    if (!frequency) continue;
    const documentFrequency = index.documentFrequency.get(term) ?? 0;
    const idf = Math.log(1 + (documents - documentFrequency + 0.5) / (documentFrequency + 0.5));
    const normalizedFrequency =
      (frequency * (k1 + 1)) / (frequency + k1 * (1 - b + (b * indexed.length) / index.averageLength));
    score += weight * idf * normalizedFrequency;
  }
  const name = normalize(indexed.tool.entry.name);
  if (name === normalizedQuery) score += 50;
  else if (name.includes(normalizedQuery)) score += 10;
  if (normalize(indexed.tool.entry.title ?? '').includes(normalizedQuery)) score += 5;
  else if (normalize(indexed.tool.entry.description).includes(normalizedQuery)) score += 3;
  return Math.round(score * 100) / 100;
}

function inputJsonSchema(tool: CapturedTool): Record<string, unknown> {
  const { $schema: _ignored, ...schema } = z.toJSONSchema(tool.schema, { target: 'draft-7' }) as Record<string, unknown>;
  return schema;
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
  private index?: ToolIndex;

  constructor(readonly config: AppConfig) {}

  capture(
    name: string,
    entry: ToolRegistryEntry,
    registration: CapturedToolRegistration,
    callback: CapturedToolCallback
  ): void {
    if (entry.category === 'gateway') return;
    this.index = undefined;
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

  search(query = '', category?: string, limit = 8, schemaCount = 3): object {
    const normalizedCategory = normalize(category ?? '');
    this.index ??= indexTools([...this.tools.values()]);
    const index = this.index;
    const terms = expandQuery(query);
    const ranked = this.visibleTools()
      .filter(tool => !normalizedCategory || normalize(tool.entry.category) === normalizedCategory)
      .map(tool => ({ tool, score: scoreTool(query, terms, index.tools.get(tool.entry.name)!, index) }))
      .filter(result => !query.trim() || result.score > 0)
      .sort((left, right) => right.score - left.score || left.tool.entry.name.localeCompare(right.tool.entry.name));
    const selected = ranked.slice(0, Math.min(Math.max(limit, 1), 20));
    return {
      query,
      category: category || undefined,
      found: selected.length,
      totalAvailable: ranked.length,
      tools: selected.map(({ tool, score }, position) => ({
        name: tool.entry.name,
        title: tool.entry.title,
        category: tool.entry.category,
        description: tool.entry.description,
        readOnly: !tool.entry.isWriteOperation,
        requiresConfirmation: tool.entry.isWriteOperation && tool.entry.requiresConfirmation ? true : undefined,
        direct: this.shouldExposeDirect(tool.entry),
        score,
        inputSchema: position < schemaCount ? inputJsonSchema(tool) : undefined
      })),
      next:
        'Tools with inputSchema can go straight to gateway_execute_tool; for the others call gateway_get_tool_schema first. Direct tools can also be called by name.'
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
      inputSchema: inputJsonSchema(tool),
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
      return errorResult('invalid_tool_arguments', validationMessage(parsed.error), {
        toolName: name,
        inputSchema: inputJsonSchema(tool)
      });
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
        'Search the signed-in user\'s authorized Microsoft 365 tools by English or Chinese intent, for anything the directly exposed tools do not cover (sending, scheduling, OneDrive and SharePoint management, contacts, users...). Top results include their input schema.',
        ['搜索可用工具', '查找能完成任务的工具', '发现更多Microsoft 365能力']
      ),
      inputSchema: {
        query: z.string().max(300).default('').describe('English or Chinese task description.'),
        category: z.string().max(50).optional().describe('Optional exact category filter, such as mail or drive.'),
        limit: z.number().int().min(1).max(20).default(8).describe('Maximum matching tools to return.'),
        schemaCount: z
          .number()
          .int()
          .min(0)
          .max(5)
          .default(3)
          .describe('Include the input schema for this many top results, so they can be executed without gateway_get_tool_schema.')
      },
      annotations: { readOnlyHint: true, openWorldHint: false }
    },
    async ({ query, category, limit, schemaCount }) => jsonResult(catalog.search(query, category, limit, schemaCount))
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
        'Execute a tool returned by gateway_search_tools with parameters matching its inputSchema. The target tool is rechecked for App Roles, confirmation policy, audit, and Graph scope controls. Invalid parameters return the expected schema.',
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
