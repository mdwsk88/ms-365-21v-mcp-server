import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { AppConfig } from './config.js';
import { registerTools } from './tools/index.js';
import type { ToolAccessContext } from './tools/types.js';

export function createGraphMcpServer(config: AppConfig, access: ToolAccessContext = {}): McpServer {
  const dynamicInstructions =
    config.toolExposureMode === 'discovery'
      ? ' Business tools are available through gateway_search_tools, gateway_get_tool_schema, and gateway_execute_tool. Search first, inspect the schema, then execute.'
      : config.toolExposureMode === 'hybrid'
        ? ' Prefer directly exposed tools. If no direct tool matches, use gateway_search_tools, gateway_get_tool_schema, and gateway_execute_tool.'
        : '';
  const server = new McpServer(
    {
      name: 'ms-365-21v-mcp-server',
      version: '0.1.0'
    },
    {
      instructions:
        'This is a 21Vianet Microsoft Graph MCP gateway for OAuth 2.1 capable clients. After sign-in, it uses on-behalf-of token exchange to call Microsoft Graph China as the signed-in user. Tool names and titles are English for client compatibility; descriptions include Chinese intent aliases. Available tools can be restricted by deployment categories and Microsoft Entra app roles. Sensitive writes may require confirmation. If a tool returns approval_url, show that URL to the user and stop. Do not substitute a draft or another write tool. Only after the user says the page was approved, call confirm_execute with confirm_token.' +
        dynamicInstructions
    }
  );

  registerTools(server, config, {
    ...access,
    bypassRoleFiltering: access.bypassRoleFiltering ?? config.transport === 'stdio'
  });
  return server;
}
