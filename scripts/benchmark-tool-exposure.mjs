#!/usr/bin/env node
import fs from 'node:fs';
import { performance } from 'node:perf_hooks';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { loadConfig } from '../dist/config.js';
import { createGraphMcpServer } from '../dist/mcp-server.js';

const example = fs.readFileSync(new URL('../.env.example', import.meta.url), 'utf8');
const disabledGraphScopes =
  example.match(/^MCP_DISABLED_GRAPH_SCOPES=(.*)$/m)?.[1].split(',').filter(Boolean) ?? [];
const directTools = [
  'graph_get_me',
  'mail_list_messages',
  'mail_search_messages',
  'calendar_list_events',
  'drive_search_items',
  'smart_mail_digest',
  'smart_calendar_conflicts'
];

async function measure(mode, roles) {
  const config = {
    ...loadConfig(),
    transport: 'http',
    inboundAuthDisabled: false,
    roleBasedFiltering: true,
    auditLogEnabled: false,
    disabledGraphScopes,
    toolExposureMode: mode,
    directToolCategories: ['smart'],
    directTools
  };
  const started = performance.now();
  const server = createGraphMcpServer(config, { userRoles: roles });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  const client = new Client({ name: 'tool-exposure-benchmark', version: '1.0.0' });
  await server.connect(serverTransport);
  await client.connect(clientTransport);
  const tools = await client.listTools();
  const elapsedMs = performance.now() - started;
  await client.close();
  await server.close();
  return {
    mode,
    roleProfile: roles.join('+'),
    tools: tools.tools.length,
    schemaBytes: Buffer.byteLength(JSON.stringify(tools)),
    localInitializeAndListMs: Number(elapsedMs.toFixed(2))
  };
}

const rows = [];
for (const roles of [['mcp.admin'], ['mcp.mail']]) {
  for (const mode of ['direct', 'discovery', 'hybrid']) {
    rows.push(await measure(mode, roles));
  }
}
console.table(rows);
console.log(
  'This local benchmark measures MCP registration and tools/list only. It does not measure model selection, network latency, OAuth, or Microsoft Graph.'
);
