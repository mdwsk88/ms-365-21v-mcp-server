import { z } from 'zod/v4';
import { getFileDownloadUrl, searchFiles } from '../graph-scenarios.js';
import { describeTool, runTool } from './results.js';
import { callerHasRoles } from './runtime.js';
import type { ToolModule } from './types.js';

export const filesModule: ToolModule = {
  category: 'drive',
  displayName: 'Files (OneDrive and SharePoint)',
  description: 'Find documents across OneDrive and SharePoint and hand agents a short-lived download URL.',
  requiredRole: 'mcp.drive',
  toolNames: ['files_search', 'files_get_download_url'],
  register(server, config) {
    server.registerTool(
      'files_search',
      {
        title: 'Search Files',
        description: describeTool(
          'Find documents across OneDrive and SharePoint sites the signed-in user can open, by keywords in name or content. Filter by site, file type and modification date. Returns location, last editor, webUrl, driveId and itemId for files_get_download_url. SharePoint results also require the mcp.sharepoint App Role.',
          ['搜索文档', '查找SharePoint文件', '查找OneDrive文件', '找某个项目的Word或PDF', '搜索网盘文件']
        ),
        inputSchema: {
          query: z.string().min(1).max(300).describe('Keywords from the file name or content.'),
          scope: z.enum(['all', 'onedrive', 'sharepoint']).optional().describe('Where to search. Defaults to all.'),
          site: z
            .string()
            .min(1)
            .max(500)
            .optional()
            .describe('Limit to one SharePoint site: its URL, for example https://tenant.sharepoint.cn/sites/Project, or site ID.'),
          fileTypes: z
            .array(z.string().min(1).max(10))
            .max(10)
            .optional()
            .describe('File extensions such as docx, pdf, pptx, xlsx.'),
          modifiedAfter: z.string().max(40).optional().describe('Only files modified at or after this ISO date or date-time.'),
          top: z.number().int().min(1).max(50).optional().describe('Maximum files to return. Defaults to 20.')
        },
        annotations: { readOnlyHint: true, openWorldHint: true }
      },
      async input =>
        runTool(async () =>
          searchFiles(config, { ...input, includeSharePoint: callerHasRoles(config, ['mcp.sharepoint']) })
        )
    );

    server.registerTool(
      'files_get_download_url',
      {
        title: 'Get File Download URL',
        description: describeTool(
          'Resolve a OneDrive or SharePoint file from a link pasted in chat or mail, or from driveId and itemId, and return a short-lived pre-authenticated download URL for the client to fetch and parse. format pdf converts Office files to PDF first. Small plain-text files also return their text inline. SharePoint files also require the mcp.sharepoint App Role.',
          ['打开SharePoint链接', '读取文档内容', '下载文件', '获取文件下载地址', '把Word转成PDF']
        ),
        inputSchema: {
          url: z
            .string()
            .min(1)
            .max(2000)
            .optional()
            .describe('OneDrive or SharePoint sharing link or file URL. Use this or itemId.'),
          driveId: z.string().min(1).max(500).optional().describe('Drive ID from files_search. Omit for the user OneDrive.'),
          itemId: z.string().min(1).max(500).optional().describe('Item ID from files_search.'),
          format: z
            .enum(['original', 'pdf'])
            .optional()
            .describe('original (default) or pdf, which converts Word, Excel and PowerPoint files to PDF.'),
          includeText: z.boolean().optional().describe('Return small plain-text files (txt, md, csv, json...) inline. Defaults to true.'),
          maxTextChars: z.number().int().min(200).max(100000).optional().describe('Maximum inline text characters. Defaults to 20000.')
        },
        annotations: { readOnlyHint: true, openWorldHint: true }
      },
      async input =>
        runTool(async () => {
          if (!input.url && !input.itemId) throw new Error('Provide url, or itemId with optional driveId.');
          return getFileDownloadUrl(config, { ...input, includeSharePoint: callerHasRoles(config, ['mcp.sharepoint']) });
        })
    );
  }
};
