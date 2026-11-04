#!/usr/bin/env node
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { loadConfig } from './config.js';
import { startHttpServer } from './http-server.js';
import { createGraphMcpServer } from './mcp-server.js';
import { formatPermissionReport, permissionReport } from './permissions-report.js';

const config = loadConfig();

if (process.argv.includes('--list-permissions')) {
  const report = permissionReport(config);
  console.log(process.argv.includes('--json') ? JSON.stringify(report, null, 2) : formatPermissionReport(report));
} else if (config.transport === 'http') {
  await startHttpServer(config);
} else {
  const server = createGraphMcpServer(config);
  const transport = new StdioServerTransport();
  await server.connect(transport);
}
