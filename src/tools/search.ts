import { z } from 'zod/v4';
import { searchMicrosoft365 } from '../graph-extended.js';
import { searchTeamsMessages } from '../graph-scenarios.js';
import { describeTool, runTool } from './results.js';
import type { ToolModule } from './types.js';

const pagingSchema = {
  from: z.number().int().min(0).max(1000).optional().describe('Zero-based result offset.'),
  size: z.number().int().min(1).max(100).optional().describe('Number of search hits to return.')
};

export const searchModule: ToolModule = {
  category: 'search',
  displayName: 'Microsoft Search',
  description: 'Cross-service Microsoft Search tools for 21V-supported entity types.',
  requiredRole: 'mcp.search',
  toolNames: [
    'search_mail',
    'search_calendar',
    'search_files',
    'search_sharepoint',
    'search_teams'
  ],
  toolMetadata: {
    search_mail: { requiredRoles: ['mcp.mail'] },
    search_calendar: { requiredRoles: ['mcp.calendar'] },
    search_files: { requiredRoles: ['mcp.drive'] },
    search_sharepoint: { requiredRoles: ['mcp.sharepoint'] },
    search_teams: { requiredRoles: ['mcp.teams'] }
  },
  register(server, config) {
    server.registerTool(
      'search_mail',
      {
        title: 'Search Microsoft 365 Mail',
        description: describeTool(
          'Search mailbox messages with Microsoft Search query syntax. Requires both mcp.search and mcp.mail roles.',
          ['全文搜索邮件', '用Microsoft Search查邮件', '跨文件夹搜索邮件']
        ),
        inputSchema: {
          query: z.string().min(1).max(1000).describe('Microsoft Search query string.'),
          ...pagingSchema
        }
      },
      async ({ query, from, size }) =>
        runTool(async () =>
          searchMicrosoft365(
            config,
            'message',
            query,
            ['Mail.Read'],
            from,
            size,
            ['id', 'subject', 'from', 'receivedDateTime', 'bodyPreview', 'webLink']
          )
        )
    );

    server.registerTool(
      'search_calendar',
      {
        title: 'Search Microsoft 365 Calendar',
        description: describeTool(
          'Search calendar events with Microsoft Search query syntax. Requires both mcp.search and mcp.calendar roles.',
          ['全文搜索日历', '按主题搜索会议', '用Microsoft Search查日程']
        ),
        inputSchema: {
          query: z.string().min(1).max(1000).describe('Microsoft Search query string.'),
          ...pagingSchema
        }
      },
      async ({ query, from, size }) =>
        runTool(async () =>
          searchMicrosoft365(
            config,
            'event',
            query,
            ['Calendars.Read'],
            from,
            size,
            ['id', 'subject', 'organizer', 'start', 'end', 'location', 'webLink']
          )
        )
    );

    server.registerTool(
      'search_files',
      {
        title: 'Search Microsoft 365 Files',
        description: describeTool(
          'Search files visible to the signed-in user across OneDrive and Microsoft 365. Requires both mcp.search and mcp.drive roles.',
          ['跨网盘搜索文件', '全文搜索Microsoft 365文件', '用Microsoft Search查文档']
        ),
        inputSchema: {
          query: z.string().min(1).max(1000).describe('Microsoft Search query string.'),
          ...pagingSchema
        }
      },
      async ({ query, from, size }) =>
        runTool(async () =>
          searchMicrosoft365(
            config,
            'driveItem',
            query,
            ['Files.Read.All'],
            from,
            size,
            ['id', 'name', 'webUrl', 'size', 'lastModifiedDateTime', 'parentReference']
          )
        )
    );

    server.registerTool(
      'search_sharepoint',
      {
        title: 'Search SharePoint Content',
        description: describeTool(
          'Search SharePoint sites or list items with Microsoft Search. Requires both mcp.search and mcp.sharepoint roles.',
          ['全文搜索SharePoint', '搜索站点内容', '跨列表搜索数据']
        ),
        inputSchema: {
          entityType: z.enum(['site', 'listItem']).describe('Search sites or SharePoint list items.'),
          query: z.string().min(1).max(1000).describe('Microsoft Search query string.'),
          ...pagingSchema
        }
      },
      async ({ entityType, query, from, size }) =>
        runTool(async () =>
          searchMicrosoft365(config, entityType, query, ['Sites.Read.All'], from, size)
        )
    );

    server.registerTool(
      'search_teams',
      {
        title: 'Search Teams Messages',
        description: describeTool(
          'Search Teams chat and channel messages by keywords with Microsoft Search. Each hit has sender, time, a highlighted summary and the chatId (or teamId and channelId) to read the conversation with teams_list_chat_messages or teams_list_channel_messages. Continue with from set to nextFrom. Requires delegated Chat.Read; channel hits also need ChannelMessage.Read.All. Requires both mcp.search and mcp.teams roles.',
          ['全文搜索Teams消息', '搜索聊天记录', '用Microsoft Search查Teams讨论', '查找谁在Teams里提到过']
        ),
        inputSchema: {
          query: z.string().min(1).max(1000).describe('Keywords or KQL, for example "budget from:Zhang".'),
          from: z.number().int().min(0).max(1000).optional().describe('Zero-based hit offset: nextFrom from the previous call.'),
          size: z.number().int().min(1).max(50).optional().describe('Hits to return, from 1 to 50. Defaults to 20.')
        },
        annotations: { readOnlyHint: true, openWorldHint: true }
      },
      async ({ query, from, size }) => runTool(async () => searchTeamsMessages(config, { query, from, size }))
    );
  }
};
