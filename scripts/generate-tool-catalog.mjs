import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

import { loadConfig } from '../dist/config.js';
import { createHydratedToolRegistry } from '../dist/tools/index.js';

const projectRoot = fileURLToPath(new URL('..', import.meta.url));
const productionDisabledGraphScopes = [
  'Channel.Create',
  'Channel.Delete.All',
  'ChannelMember.Read.All',
  'ChannelMember.ReadWrite.All',
  'ChannelMessage.Read.All',
  'ChannelMessage.ReadWrite',
  'ChannelMessage.Send',
  'ChannelSettings.ReadWrite.All',
  'ChatMember.ReadWrite',
  'GroupMember.Read.All',
  'TeamMember.Read.All',
  'TeamMember.ReadWrite.All',
  'User.Read.All'
];
const baseConfig = {
  ...loadConfig(),
  transport: 'http',
  toolCategories: [],
  enableRawGraphGet: false,
  roleBasedFiltering: true,
  auditLogEnabled: false,
  confirmOperations: ['production-confirmation-enabled']
};
const implementationRegistry = createHydratedToolRegistry({ ...baseConfig, disabledGraphScopes: [] });
const registry = createHydratedToolRegistry({
  ...baseConfig,
  disabledGraphScopes: productionDisabledGraphScopes
});
const entries = registry.getAll();
const enabledNames = new Set(entries.map(entry => entry.name));
const disabledEntries = implementationRegistry.getAll().filter(entry => !enabledNames.has(entry.name));

const lines = [
  '# MS 365-21V MCP Server Pro - Tool Catalog',
  '',
  '> This file is generated from the runtime registry by `npm run docs:tools`. Do not edit tool rows manually.',
  '',
  `The implementation contains **${implementationRegistry.getAll().length} tools**. The current production permission boundary exposes **${entries.length} tools** and hides **${disabledEntries.length} tools** whose Graph delegated scopes are not approved. Every enabled business tool is mapped to an Entra App Role and one or more Microsoft Graph delegated permissions.`,
  '',
  '## Summary',
  '',
  '| Module | Tools | Read | Write | App Role |',
  '|---|---:|---:|---:|---|'
];

for (const category of registry.getAllCategories()) {
  const tools = registry.getByCategory(category);
  const writes = tools.filter((tool) => tool.isWriteOperation).length;
  const roles = unique(tools.flatMap((tool) => tool.requiredRoles ?? [])).map(code).join(', ') || '-';
  lines.push(`| ${escapeCell(tools[0]?.module ?? category)} | ${tools.length} | ${tools.length - writes} | ${writes} | ${roles} |`);
}

for (const category of registry.getAllCategories()) {
  const tools = registry.getByCategory(category).sort((left, right) => left.name.localeCompare(right.name));
  lines.push('', `## ${tools[0]?.module ?? category}`, '', '| Tool | Title | Access | Required App Roles | Graph delegated scopes |', '|---|---|---|---|---|');
  for (const tool of tools) {
    const roles = (tool.requiredRoles ?? []).map(code).join(', ') || '-';
    const scopes = tool.graphScopes.map(code).join(', ') || '-';
    lines.push(`| ${code(tool.name)} | ${escapeCell(tool.title ?? tool.name)} | ${tool.isWriteOperation ? 'Write' : 'Read'} | ${roles} | ${scopes} |`);
  }
}

lines.push(
  '',
  '## Disabled Pending Tenant Approval',
  '',
  `The following **${disabledEntries.length} tools** remain implemented but are excluded from both MCP discovery and direct invocation. They become available after the corresponding delegated permissions are approved and removed from \`MCP_DISABLED_GRAPH_SCOPES\`.`,
  '',
  '| Tool | Module | Graph delegated scopes |',
  '|---|---|---|',
  ...disabledEntries
    .sort((left, right) => left.name.localeCompare(right.name))
    .map(tool => `| ${code(tool.name)} | ${escapeCell(tool.module)} | ${tool.graphScopes.map(code).join(', ')} |`),
  '',
  '## Coverage Boundary',
  '',
  'The catalog targets Microsoft Graph v1.0 APIs that are available in the China cloud operated by 21Vianet, support delegated work-or-school identities, and are appropriate for interactive MCP use. It intentionally excludes beta-only APIs, application-only tenant administration, subscriptions/webhooks, security and Intune administration, and workloads whose Microsoft documentation marks China as unsupported.'
);

await fs.writeFile(path.join(projectRoot, 'docs', 'TOOL_CATALOG.md'), `${lines.join('\n')}\n`, 'utf8');

function unique(values) {
  return [...new Set(values)].sort();
}

function code(value) {
  return `\`${String(value).replace(/`/g, '\\`')}\``;
}

function escapeCell(value) {
  return String(value).replace(/\|/g, '\\|').replace(/\r?\n/g, ' ');
}
