import { z } from 'zod/v4';
import { graphGet } from '../graph.js';
import { describeTool, runTool } from './results.js';
import type { ToolModule } from './types.js';

export const debugModule: ToolModule = {
  category: 'debug',
  displayName: 'Graph Debug',
  description: 'Development-only raw Graph GET helper.',
  requiredRole: 'mcp.admin',
  isEnabled: (config) => config.enableRawGraphGet,
  toolNames: ['graph_debug_get'],
  register(server, config) {
    if (config.enableRawGraphGet) {
      server.registerTool(
        'graph_debug_get',
        {
          title: 'Debug Graph GET Request',
          description: describeTool(
            'Development-only read-only Microsoft Graph GET helper. It only allows URLs under MS_GRAPH_BASE_URL and is disabled unless MS_ENABLE_RAW_GRAPH_GET=true.',
            ['调试Graph请求', '读取Graph接口', '开发调试查询']
          ),
          inputSchema: {
            pathOrUrl: z
              .string()
              .min(1)
              .max(2000)
              .describe(
                'Graph path or full URL, for example /me/memberOf or https://microsoftgraph.chinacloudapi.cn/v1.0/me.'
              ),
            top: z
              .number()
              .int()
              .min(1)
              .max(100)
              .optional()
              .describe('Optional $top query parameter for collection endpoints.')
          }
        },
        async ({ pathOrUrl, top }) => runTool(async () => graphGet(config, pathOrUrl, top ? { $top: top } : undefined))
      );
    }
  }
};
