#!/usr/bin/env node
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { loadConfig } from './config.js';
import { startHttpServer } from './http-server.js';
import { createGraphMcpServer } from './mcp-server.js';

const config = loadConfig();

if (config.transport === 'http') {
  await startHttpServer(config);
} else {
  const server = createGraphMcpServer(config);
  const transport = new StdioServerTransport();
  await server.connect(transport);
}
