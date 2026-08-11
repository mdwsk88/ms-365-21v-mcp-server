import { z } from 'zod/v4';
import {
  copySharePointDriveItem,
  createSharePointColumn,
  createSharePointDriveItemLink,
  createSharePointList,
  deleteSharePointColumn,
  deleteSharePointList,
  downloadSharePointFile,
  getSharePointListItemDelta,
  getSharePointSiteByPath,
  inviteSharePointDriveItem,
  listSharePointColumns,
  listSharePointDriveItemPermissions,
  listSharePointDriveVersions,
  moveSharePointDriveItem,
  renameSharePointDriveItem,
  restoreSharePointDriveVersion,
  updateSharePointColumn,
  updateSharePointList
} from '../graph-extended.js';
import {
  createSharePointDriveFolder,
  createSharePointListItem,
  deleteSharePointDriveItem,
  deleteSharePointListItem,
  getSharePointDriveItem,
  getSharePointListItem,
  getSharePointSite,
  listSharePointDriveItems,
  listSharePointDrives,
  listSharePointListItems,
  listSharePointLists,
  searchSharePointDriveItems,
  searchSharePointSites,
  updateSharePointListItem,
  uploadSmallSharePointFile
} from '../graph.js';
import { sharePointFieldNameSchema, sharePointFieldsSchema } from './schemas.js';
import { describeTool, runTool } from './results.js';
import type { ToolModule } from './types.js';

const columnDefinitionSchema = z
  .record(z.string().min(1).max(100), z.unknown())
  .refine((value) => Object.keys(value).length > 0, 'A SharePoint column definition is required.');

export const sharepointModule: ToolModule = {
  category: 'sharepoint',
  displayName: 'SharePoint',
  description: 'SharePoint sites, lists, and document library tools.',
  requiredRole: 'mcp.sharepoint',
  toolNames: [
    'sharepoint_search_sites',
    'sharepoint_get_site',
    'sharepoint_get_site_by_path',
    'sharepoint_list_lists',
    'sharepoint_create_list',
    'sharepoint_update_list',
    'sharepoint_delete_list',
    'sharepoint_list_columns',
    'sharepoint_create_column',
    'sharepoint_update_column',
    'sharepoint_delete_column',
    'sharepoint_list_list_items',
    'sharepoint_list_item_delta',
    'sharepoint_get_list_item',
    'sharepoint_create_list_item',
    'sharepoint_update_list_item',
    'sharepoint_delete_list_item',
    'sharepoint_list_drives',
    'sharepoint_list_drive_items',
    'sharepoint_search_drive_items',
    'sharepoint_get_drive_item',
    'sharepoint_download_file',
    'sharepoint_create_drive_folder',
    'sharepoint_upload_small_file',
    'sharepoint_rename_drive_item',
    'sharepoint_move_drive_item',
    'sharepoint_copy_drive_item',
    'sharepoint_create_share_link',
    'sharepoint_invite_drive_item',
    'sharepoint_list_drive_permissions',
    'sharepoint_list_drive_versions',
    'sharepoint_restore_drive_version',
    'sharepoint_delete_drive_item'
  ],
  register(server, config) {
    server.registerTool(
      'sharepoint_search_sites',
      {
        title: 'Search SharePoint Sites',
        description: describeTool(
          'Search SharePoint sites visible to the signed-in user. Requires delegated Sites.Read.All.',
          ['搜索SharePoint站点', '查找公司站点', '查找团队网站', '查看SharePoint网站']
        ),
        inputSchema: {
          query: z
            .string()
            .min(1)
            .max(200)
            .describe('Site search keywords, such as department, project, or site name.'),
          top: z.number().int().min(1).max(50).optional().describe('Number of sites to return, from 1 to 50.')
        }
      },
      async ({ query, top }) => runTool(async () => searchSharePointSites(config, query, top))
    );

    server.registerTool(
      'sharepoint_get_site',
      {
        title: 'Get SharePoint Site',
        description: describeTool('Read SharePoint site metadata by site ID. Requires delegated Sites.Read.All.', [
          '查看SharePoint站点详情',
          '读取站点信息',
          '查看站点地址'
        ]),
        inputSchema: {
          siteId: z.string().min(1).max(500).describe('SharePoint site ID, usually from sharepoint_search_sites.')
        }
      },
      async ({ siteId }) => runTool(async () => getSharePointSite(config, siteId))
    );

    server.registerTool(
      'sharepoint_list_lists',
      {
        title: 'List SharePoint Lists',
        description: describeTool(
          'List SharePoint lists and document-library lists in a site. Requires delegated Sites.Read.All.',
          ['查看SharePoint列表', '列出站点Lists', '查看站点文档库', '查看站点中的清单']
        ),
        inputSchema: {
          siteId: z.string().min(1).max(500).describe('SharePoint site ID.'),
          top: z.number().int().min(1).max(100).optional().describe('Number of lists to return, from 1 to 100.')
        }
      },
      async ({ siteId, top }) => runTool(async () => listSharePointLists(config, siteId, top))
    );

    server.registerTool(
      'sharepoint_list_list_items',
      {
        title: 'List SharePoint List Items',
        description: describeTool(
          'List items and field values in a SharePoint list. Optional filters work best on indexed columns. Requires delegated Sites.Read.All.',
          ['查看SharePoint列表项', '读取SharePoint清单数据', '查询列表记录', '筛选SharePoint列表']
        ),
        inputSchema: {
          siteId: z.string().min(1).max(500).describe('SharePoint site ID.'),
          listId: z.string().min(1).max(300).describe('SharePoint list ID from sharepoint_list_lists.'),
          top: z.number().int().min(1).max(100).optional().describe('Number of list items to return, from 1 to 100.'),
          fieldNames: z
            .array(sharePointFieldNameSchema)
            .max(30)
            .optional()
            .describe('Optional SharePoint internal column names to return. Omit to return all available fields.'),
          filter: z
            .string()
            .min(1)
            .max(1000)
            .optional()
            .describe("Optional OData filter, for example fields/Status eq 'Open'. Prefer indexed columns.")
        }
      },
      async ({ siteId, listId, top, fieldNames, filter }) =>
        runTool(async () => listSharePointListItems(config, siteId, listId, top, fieldNames, filter))
    );

    server.registerTool(
      'sharepoint_get_list_item',
      {
        title: 'Get SharePoint List Item',
        description: describeTool('Read one SharePoint list item and its fields. Requires delegated Sites.Read.All.', [
          '查看SharePoint列表项详情',
          '读取一条清单记录',
          '查看列表字段'
        ]),
        inputSchema: {
          siteId: z.string().min(1).max(500).describe('SharePoint site ID.'),
          listId: z.string().min(1).max(300).describe('SharePoint list ID.'),
          itemId: z.string().min(1).max(300).describe('SharePoint list item ID.'),
          fieldNames: z
            .array(sharePointFieldNameSchema)
            .max(30)
            .optional()
            .describe('Optional SharePoint internal column names to return.')
        }
      },
      async ({ siteId, listId, itemId, fieldNames }) =>
        runTool(async () => getSharePointListItem(config, siteId, listId, itemId, fieldNames))
    );

    server.registerTool(
      'sharepoint_create_list_item',
      {
        title: 'Create SharePoint List Item',
        description: describeTool(
          'Create an item in a SharePoint list using internal column names. Requires delegated Sites.ReadWrite.All.',
          ['创建SharePoint列表项', '新增清单记录', '向SharePoint列表写入数据']
        ),
        inputSchema: {
          siteId: z.string().min(1).max(500).describe('SharePoint site ID.'),
          listId: z.string().min(1).max(300).describe('SharePoint list ID.'),
          fields: sharePointFieldsSchema.describe(
            'Field values keyed by SharePoint internal column name, for example {"Title":"Demo"}.'
          )
        }
      },
      async ({ siteId, listId, fields }) =>
        runTool(async () => createSharePointListItem(config, siteId, listId, fields))
    );

    server.registerTool(
      'sharepoint_update_list_item',
      {
        title: 'Update SharePoint List Item',
        description: describeTool(
          'Update selected fields on a SharePoint list item. Unspecified fields remain unchanged. Requires delegated Sites.ReadWrite.All.',
          ['更新SharePoint列表项', '修改清单记录', '编辑SharePoint字段']
        ),
        inputSchema: {
          siteId: z.string().min(1).max(500).describe('SharePoint site ID.'),
          listId: z.string().min(1).max(300).describe('SharePoint list ID.'),
          itemId: z.string().min(1).max(300).describe('SharePoint list item ID.'),
          fields: sharePointFieldsSchema.describe(
            'Only the SharePoint fields to update, keyed by internal column name.'
          )
        }
      },
      async ({ siteId, listId, itemId, fields }) =>
        runTool(async () => updateSharePointListItem(config, siteId, listId, itemId, fields))
    );

    server.registerTool(
      'sharepoint_delete_list_item',
      {
        title: 'Delete SharePoint List Item',
        description: describeTool(
          'Delete an item from a SharePoint list. This is a write operation and may move the item to the site recycle bin. Requires delegated Sites.ReadWrite.All.',
          ['删除SharePoint列表项', '删除清单记录', '移除SharePoint数据']
        ),
        inputSchema: {
          siteId: z.string().min(1).max(500).describe('SharePoint site ID.'),
          listId: z.string().min(1).max(300).describe('SharePoint list ID.'),
          itemId: z.string().min(1).max(300).describe('SharePoint list item ID to delete.')
        }
      },
      async ({ siteId, listId, itemId }) =>
        runTool(async () => deleteSharePointListItem(config, siteId, listId, itemId))
    );

    server.registerTool(
      'sharepoint_list_drives',
      {
        title: 'List SharePoint Document Libraries',
        description: describeTool(
          'List document libraries available in a SharePoint site. Each library is returned as a drive. Requires delegated Files.Read.All.',
          ['查看SharePoint文档库', '列出站点文件库', '查看团队文件库']
        ),
        inputSchema: {
          siteId: z.string().min(1).max(500).describe('SharePoint site ID.'),
          top: z
            .number()
            .int()
            .min(1)
            .max(100)
            .optional()
            .describe('Number of document libraries to return, from 1 to 100.')
        }
      },
      async ({ siteId, top }) => runTool(async () => listSharePointDrives(config, siteId, top))
    );

    server.registerTool(
      'sharepoint_list_drive_items',
      {
        title: 'List SharePoint Library Items',
        description: describeTool(
          'List files and folders in a SharePoint document library root or folder. Requires delegated Files.Read.All.',
          ['查看SharePoint文件', '列出文档库文件', '查看SharePoint文件夹', '浏览团队文件']
        ),
        inputSchema: {
          driveId: z
            .string()
            .min(1)
            .max(500)
            .describe('SharePoint document-library drive ID from sharepoint_list_drives.'),
          parentItemId: z
            .string()
            .min(1)
            .max(500)
            .optional()
            .describe('Optional folder item ID. Defaults to the library root.'),
          top: z
            .number()
            .int()
            .min(1)
            .max(100)
            .optional()
            .describe('Number of files or folders to return, from 1 to 100.')
        }
      },
      async ({ driveId, parentItemId, top }) =>
        runTool(async () => listSharePointDriveItems(config, driveId, parentItemId, top))
    );

    server.registerTool(
      'sharepoint_search_drive_items',
      {
        title: 'Search SharePoint Library',
        description: describeTool(
          'Search files and folders in a SharePoint document library. Requires delegated Files.Read.All.',
          ['搜索SharePoint文件', '在文档库查文件', '查找团队文档']
        ),
        inputSchema: {
          driveId: z.string().min(1).max(500).describe('SharePoint document-library drive ID.'),
          query: z.string().min(1).max(200).describe('File or folder search keywords.'),
          top: z.number().int().min(1).max(50).optional().describe('Number of search results to return, from 1 to 50.')
        }
      },
      async ({ driveId, query, top }) => runTool(async () => searchSharePointDriveItems(config, driveId, query, top))
    );

    server.registerTool(
      'sharepoint_get_drive_item',
      {
        title: 'Get SharePoint Library Item',
        description: describeTool(
          'Read metadata for a SharePoint document-library file or folder. Requires delegated Files.Read.All.',
          ['查看SharePoint文件详情', '读取文档库文件信息', '查看SharePoint文件元数据']
        ),
        inputSchema: {
          driveId: z.string().min(1).max(500).describe('SharePoint document-library drive ID.'),
          itemId: z.string().min(1).max(500).describe('File or folder item ID.')
        }
      },
      async ({ driveId, itemId }) => runTool(async () => getSharePointDriveItem(config, driveId, itemId))
    );

    server.registerTool(
      'sharepoint_create_drive_folder',
      {
        title: 'Create SharePoint Library Folder',
        description: describeTool(
          'Create a folder in a SharePoint document library root or parent folder. Requires delegated Files.ReadWrite.All.',
          ['创建SharePoint文件夹', '在文档库新建目录', '新建团队文件夹']
        ),
        inputSchema: {
          driveId: z.string().min(1).max(500).describe('SharePoint document-library drive ID.'),
          name: z.string().min(1).max(255).describe('New folder name.'),
          parentItemId: z
            .string()
            .min(1)
            .max(500)
            .optional()
            .describe('Optional parent folder item ID. Defaults to the library root.'),
          conflictBehavior: z
            .enum(['rename', 'replace', 'fail'])
            .optional()
            .describe('Name conflict behavior. Defaults to rename.')
        }
      },
      async ({ driveId, name, parentItemId, conflictBehavior }) =>
        runTool(async () => createSharePointDriveFolder(config, driveId, name, parentItemId, conflictBehavior))
    );

    server.registerTool(
      'sharepoint_upload_small_file',
      {
        title: 'Upload Small SharePoint File',
        description: describeTool(
          'Upload a small file to a SharePoint document library using a path relative to the library root. Requires delegated Files.ReadWrite.All.',
          ['上传SharePoint文件', '上传文档到站点', '保存文件到团队文档库']
        ),
        inputSchema: {
          driveId: z.string().min(1).max(500).describe('SharePoint document-library drive ID.'),
          path: z
            .string()
            .min(1)
            .max(1000)
            .describe('Target path relative to the library root, for example Reports/demo.txt.'),
          content: z.string().min(1).describe('File content. Use raw text for text mode or base64 for base64 mode.'),
          contentEncoding: z.enum(['text', 'base64']).optional().describe('Content encoding. Defaults to text.'),
          contentType: z.string().min(1).max(200).optional().describe('Optional MIME type.'),
          conflictBehavior: z
            .enum(['replace', 'rename', 'fail'])
            .optional()
            .describe('Name conflict behavior. Defaults to replace.')
        }
      },
      async ({ driveId, path, content, contentEncoding, contentType, conflictBehavior }) =>
        runTool(async () =>
          uploadSmallSharePointFile(config, driveId, path, content, contentEncoding, contentType, conflictBehavior)
        )
    );

    server.registerTool(
      'sharepoint_delete_drive_item',
      {
        title: 'Delete SharePoint Library Item',
        description: describeTool(
          'Delete a file or folder from a SharePoint document library. This is a write operation and may move the item to the recycle bin. Requires delegated Files.ReadWrite.All.',
          ['删除SharePoint文件', '删除文档库文件', '删除SharePoint文件夹']
        ),
        inputSchema: {
          driveId: z.string().min(1).max(500).describe('SharePoint document-library drive ID.'),
          itemId: z.string().min(1).max(500).describe('File or folder item ID to delete.')
        }
      },
      async ({ driveId, itemId }) => runTool(async () => deleteSharePointDriveItem(config, driveId, itemId))
    );

    server.registerTool(
      'sharepoint_get_site_by_path',
      {
        title: 'Get SharePoint Site By Path',
        description: describeTool('Resolve a SharePoint site by hostname and server-relative path.', [
          '按网址查SharePoint站点',
          '通过站点路径获取Site ID',
          '解析SharePoint链接'
        ]),
        inputSchema: {
          hostname: z
            .string()
            .min(1)
            .max(255)
            .regex(/^[A-Za-z0-9.-]+$/, 'Use a SharePoint hostname such as contoso.sharepoint.cn.'),
          relativePath: z.string().min(1).max(1000).describe('Server-relative path, such as /sites/ProjectA.')
        }
      },
      async ({ hostname, relativePath }) =>
        runTool(async () => getSharePointSiteByPath(config, hostname, relativePath))
    );

    server.registerTool(
      'sharepoint_create_list',
      {
        title: 'Create SharePoint List',
        description: describeTool(
          'Create a SharePoint list. This administrative write requires delegated Sites.Manage.All.',
          ['创建SharePoint列表', '新建项目清单', '创建站点List']
        ),
        inputSchema: {
          siteId: z.string().min(1).max(500).describe('SharePoint site ID.'),
          displayName: z.string().min(1).max(255).describe('List display name.'),
          description: z.string().max(1000).optional().describe('Optional list description.'),
          template: z
            .enum(['genericList', 'documentLibrary', 'survey', 'links', 'announcements', 'contacts'])
            .optional()
            .describe('List template. Defaults to genericList.'),
          columns: z
            .array(columnDefinitionSchema)
            .max(30)
            .optional()
            .describe('Optional Microsoft Graph columnDefinition objects.')
        }
      },
      async ({ siteId, displayName, description, template, columns }) =>
        runTool(async () => createSharePointList(config, siteId, displayName, description, template, columns))
    );

    server.registerTool(
      'sharepoint_update_list',
      {
        title: 'Update SharePoint List',
        description: describeTool('Update the name or description of a SharePoint list. Requires Sites.Manage.All.', [
          '修改SharePoint列表',
          '重命名站点清单',
          '更新列表描述'
        ]),
        inputSchema: {
          siteId: z.string().min(1).max(500).describe('SharePoint site ID.'),
          listId: z.string().min(1).max(300).describe('SharePoint list ID.'),
          displayName: z.string().min(1).max(255).optional().describe('New display name.'),
          description: z.string().max(1000).optional().describe('New description.')
        }
      },
      async ({ siteId, listId, displayName, description }) =>
        runTool(async () => updateSharePointList(config, siteId, listId, { displayName, description }))
    );

    server.registerTool(
      'sharepoint_delete_list',
      {
        title: 'Delete SharePoint List',
        description: describeTool('Delete an entire SharePoint list. Requires delegated Sites.Manage.All.', [
          '删除SharePoint列表',
          '删除站点清单',
          '移除整个List'
        ]),
        inputSchema: {
          siteId: z.string().min(1).max(500).describe('SharePoint site ID.'),
          listId: z.string().min(1).max(300).describe('List ID to delete.')
        }
      },
      async ({ siteId, listId }) => runTool(async () => deleteSharePointList(config, siteId, listId))
    );

    server.registerTool(
      'sharepoint_list_columns',
      {
        title: 'List SharePoint List Columns',
        description: describeTool('List column definitions and internal names for a SharePoint list.', [
          '查看SharePoint列',
          '查看列表字段定义',
          '获取列内部名称'
        ]),
        inputSchema: {
          siteId: z.string().min(1).max(500).describe('SharePoint site ID.'),
          listId: z.string().min(1).max(300).describe('SharePoint list ID.')
        }
      },
      async ({ siteId, listId }) => runTool(async () => listSharePointColumns(config, siteId, listId))
    );

    server.registerTool(
      'sharepoint_create_column',
      {
        title: 'Create SharePoint List Column',
        description: describeTool(
          'Create a SharePoint list column from a Microsoft Graph columnDefinition. Requires Sites.Manage.All.',
          ['创建SharePoint列', '给列表添加字段', '新增清单列']
        ),
        inputSchema: {
          siteId: z.string().min(1).max(500).describe('SharePoint site ID.'),
          listId: z.string().min(1).max(300).describe('SharePoint list ID.'),
          definition: columnDefinitionSchema.describe(
            'Graph columnDefinition, for example {"name":"Status","choice":{"choices":["Open","Closed"]}}.'
          )
        }
      },
      async ({ siteId, listId, definition }) =>
        runTool(async () => createSharePointColumn(config, siteId, listId, definition))
    );

    server.registerTool(
      'sharepoint_update_column',
      {
        title: 'Update SharePoint List Column',
        description: describeTool('Update writable properties on a SharePoint list column. Requires Sites.Manage.All.', [
          '修改SharePoint列',
          '更新列表字段',
          '调整清单列配置'
        ]),
        inputSchema: {
          siteId: z.string().min(1).max(500).describe('SharePoint site ID.'),
          listId: z.string().min(1).max(300).describe('SharePoint list ID.'),
          columnId: z.string().min(1).max(300).describe('Column ID.'),
          definition: columnDefinitionSchema.describe('Only the column properties to update.')
        }
      },
      async ({ siteId, listId, columnId, definition }) =>
        runTool(async () => updateSharePointColumn(config, siteId, listId, columnId, definition))
    );

    server.registerTool(
      'sharepoint_delete_column',
      {
        title: 'Delete SharePoint List Column',
        description: describeTool('Delete a custom SharePoint list column. Requires Sites.Manage.All.', [
          '删除SharePoint列',
          '移除列表字段',
          '删除清单列'
        ]),
        inputSchema: {
          siteId: z.string().min(1).max(500).describe('SharePoint site ID.'),
          listId: z.string().min(1).max(300).describe('SharePoint list ID.'),
          columnId: z.string().min(1).max(300).describe('Column ID to delete.')
        }
      },
      async ({ siteId, listId, columnId }) =>
        runTool(async () => deleteSharePointColumn(config, siteId, listId, columnId))
    );

    server.registerTool(
      'sharepoint_list_item_delta',
      {
        title: 'Read SharePoint List Changes',
        description: describeTool(
          'Read newly created, updated, or deleted SharePoint list items and return nextLink or deltaLink state.',
          ['查看SharePoint列表变更', '增量同步清单', '获取列表Delta']
        ),
        inputSchema: {
          siteId: z.string().min(1).max(500).describe('SharePoint site ID.'),
          listId: z.string().min(1).max(300).describe('SharePoint list ID.'),
          token: z
            .string()
            .min(1)
            .max(4000)
            .optional()
            .describe('Optional delta token, or latest to request a fresh delta link.'),
          top: z.number().int().min(1).max(100).optional().describe('Page size.')
        }
      },
      async ({ siteId, listId, token, top }) =>
        runTool(async () => getSharePointListItemDelta(config, siteId, listId, token, top))
    );

    server.registerTool(
      'sharepoint_download_file',
      {
        title: 'Download SharePoint File',
        description: describeTool(
          'Download a document-library file as base64. The MCP response is limited to 10 MB.',
          ['下载SharePoint文件', '读取文档库文件内容', '获取团队文件base64']
        ),
        inputSchema: {
          driveId: z.string().min(1).max(500).describe('Document-library drive ID.'),
          itemId: z.string().min(1).max(500).describe('File item ID.')
        }
      },
      async ({ driveId, itemId }) => runTool(async () => downloadSharePointFile(config, driveId, itemId))
    );

    server.registerTool(
      'sharepoint_rename_drive_item',
      {
        title: 'Rename SharePoint Library Item',
        description: describeTool('Rename a file or folder in a SharePoint document library.', [
          '重命名SharePoint文件',
          '修改文档库文件名',
          '重命名团队文件夹'
        ]),
        inputSchema: {
          driveId: z.string().min(1).max(500).describe('Document-library drive ID.'),
          itemId: z.string().min(1).max(500).describe('File or folder item ID.'),
          newName: z.string().min(1).max(255).describe('New name.')
        }
      },
      async ({ driveId, itemId, newName }) =>
        runTool(async () => renameSharePointDriveItem(config, driveId, itemId, newName))
    );

    server.registerTool(
      'sharepoint_move_drive_item',
      {
        title: 'Move SharePoint Library Item',
        description: describeTool('Move a file or folder to another folder in the same document library.', [
          '移动SharePoint文件',
          '把文档移到文件夹',
          '整理团队文件'
        ]),
        inputSchema: {
          driveId: z.string().min(1).max(500).describe('Document-library drive ID.'),
          itemId: z.string().min(1).max(500).describe('Source item ID.'),
          newParentItemId: z.string().min(1).max(500).describe('Destination parent folder ID.')
        }
      },
      async ({ driveId, itemId, newParentItemId }) =>
        runTool(async () => moveSharePointDriveItem(config, driveId, itemId, newParentItemId))
    );

    server.registerTool(
      'sharepoint_copy_drive_item',
      {
        title: 'Copy SharePoint Library Item',
        description: describeTool('Copy a SharePoint file or folder asynchronously.', [
          '复制SharePoint文件',
          '复制文档库文件夹',
          '创建团队文件副本'
        ]),
        inputSchema: {
          driveId: z.string().min(1).max(500).describe('Document-library drive ID.'),
          itemId: z.string().min(1).max(500).describe('Source item ID.'),
          parentItemId: z.string().min(1).max(500).describe('Destination parent folder ID.'),
          newName: z.string().min(1).max(255).optional().describe('Optional copy name.')
        }
      },
      async ({ driveId, itemId, parentItemId, newName }) =>
        runTool(async () => copySharePointDriveItem(config, driveId, itemId, parentItemId, newName))
    );

    server.registerTool(
      'sharepoint_create_share_link',
      {
        title: 'Create SharePoint Share Link',
        description: describeTool('Create a sharing link for a SharePoint file or folder.', [
          '创建SharePoint分享链接',
          '分享团队文件',
          '生成文档库链接'
        ]),
        inputSchema: {
          driveId: z.string().min(1).max(500).describe('Document-library drive ID.'),
          itemId: z.string().min(1).max(500).describe('File or folder item ID.'),
          type: z.enum(['view', 'edit', 'embed']).optional().describe('Link type. Defaults to view.'),
          scope: z.enum(['anonymous', 'organization', 'users']).optional().describe('Sharing scope.')
        }
      },
      async ({ driveId, itemId, type, scope }) =>
        runTool(async () => createSharePointDriveItemLink(config, driveId, itemId, type, scope))
    );

    server.registerTool(
      'sharepoint_invite_drive_item',
      {
        title: 'Invite People To SharePoint Item',
        description: describeTool('Grant selected people read or write access to a SharePoint file or folder.', [
          '共享SharePoint文件给指定人员',
          '邀请同事访问文档库',
          '授予团队文件权限'
        ]),
        inputSchema: {
          driveId: z.string().min(1).max(500).describe('Document-library drive ID.'),
          itemId: z.string().min(1).max(500).describe('File or folder item ID.'),
          recipients: z.array(z.object({
            email: z.string().email().describe('Recipient email address.'),
            name: z.string().min(1).max(120).optional().describe('Optional display name.')
          })).min(1).max(20).describe('People to invite.'),
          roles: z.array(z.enum(['read', 'write'])).min(1).max(2).describe('Granted roles.'),
          message: z.string().max(2000).optional().describe('Optional invitation message.'),
          requireSignIn: z.boolean().optional().describe('Require sign-in. Defaults to true.'),
          sendInvitation: z.boolean().optional().describe('Send invitation email. Defaults to true.')
        }
      },
      async ({ driveId, itemId, recipients, roles, message, requireSignIn, sendInvitation }) =>
        runTool(async () =>
          inviteSharePointDriveItem(
            config,
            driveId,
            itemId,
            recipients,
            roles,
            message,
            requireSignIn,
            sendInvitation
          )
        )
    );

    server.registerTool(
      'sharepoint_list_drive_permissions',
      {
        title: 'List SharePoint Item Permissions',
        description: describeTool('List sharing and permission entries for a SharePoint file or folder.', [
          '查看SharePoint文件权限',
          '查看谁能访问团队文件',
          '列出文档库共享权限'
        ]),
        inputSchema: {
          driveId: z.string().min(1).max(500).describe('Document-library drive ID.'),
          itemId: z.string().min(1).max(500).describe('File or folder item ID.')
        }
      },
      async ({ driveId, itemId }) =>
        runTool(async () => listSharePointDriveItemPermissions(config, driveId, itemId))
    );

    server.registerTool(
      'sharepoint_list_drive_versions',
      {
        title: 'List SharePoint File Versions',
        description: describeTool('List retained versions of a SharePoint document-library file.', [
          '查看SharePoint版本历史',
          '列出团队文档旧版本',
          '查看文档库版本'
        ]),
        inputSchema: {
          driveId: z.string().min(1).max(500).describe('Document-library drive ID.'),
          itemId: z.string().min(1).max(500).describe('File item ID.')
        }
      },
      async ({ driveId, itemId }) => runTool(async () => listSharePointDriveVersions(config, driveId, itemId))
    );

    server.registerTool(
      'sharepoint_restore_drive_version',
      {
        title: 'Restore SharePoint File Version',
        description: describeTool('Restore a historical version of a SharePoint document-library file.', [
          '恢复SharePoint文件版本',
          '回滚团队文档',
          '还原文档库旧版本'
        ]),
        inputSchema: {
          driveId: z.string().min(1).max(500).describe('Document-library drive ID.'),
          itemId: z.string().min(1).max(500).describe('File item ID.'),
          versionId: z.string().min(1).max(300).describe('Version ID.')
        }
      },
      async ({ driveId, itemId, versionId }) =>
        runTool(async () => restoreSharePointDriveVersion(config, driveId, itemId, versionId))
    );
  }
};
