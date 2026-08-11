import { z } from 'zod/v4';
import {
  copyDriveItem,
  downloadDriveFile,
  getMyDrive,
  inviteDriveItem,
  listDriveChildren,
  listDriveVersions,
  listRecentDriveItems,
  restoreDriveVersion
} from '../graph-extended.js';
import {
  createDriveFolder,
  createDriveItemLink,
  deleteDriveItem,
  getDriveItem,
  listDriveItemPermissions,
  listMyDriveRoot,
  moveDriveItem,
  renameDriveItem,
  searchMyDrive,
  uploadSmallDriveFile
} from '../graph.js';
import { describeTool, runTool } from './results.js';
import type { ToolModule } from './types.js';

export const driveModule: ToolModule = {
  category: 'drive',
  displayName: 'OneDrive',
  description: 'OneDrive file and folder tools.',
  requiredRole: 'mcp.drive',
  toolNames: [
    'drive_get_drive',
    'drive_list_root',
    'drive_list_children',
    'drive_list_recent',
    'drive_search_items',
    'drive_get_item',
    'drive_download_file',
    'drive_create_folder',
    'drive_upload_small_file',
    'drive_rename_item',
    'drive_move_item',
    'drive_copy_item',
    'drive_delete_item',
    'drive_create_share_link',
    'drive_invite_item',
    'drive_list_permissions',
    'drive_list_versions',
    'drive_restore_version'
  ],
  register(server, config) {
    server.registerTool(
      'drive_get_drive',
      {
        title: 'Get My OneDrive',
        description: describeTool('Read signed-in user OneDrive metadata, owner, type, and quota.', [
          '查看我的OneDrive信息',
          '查看网盘容量',
          '查看OneDrive配额'
        ])
      },
      async () => runTool(async () => getMyDrive(config))
    );

    server.registerTool(
      'drive_list_root',
      {
        title: 'List My OneDrive Root',
        description: describeTool(
          'List files and folders in the signed-in user OneDrive root. Requires delegated Files.Read.',
          ['查看我的OneDrive', '查看网盘根目录', '列出OneDrive文件']
        ),
        inputSchema: {
          top: z
            .number()
            .int()
            .min(1)
            .max(100)
            .optional()
            .describe('Number of files or folders to return, from 1 to 100.')
        }
      },
      async ({ top }) => runTool(async () => listMyDriveRoot(config, top))
    );

    server.registerTool(
      'drive_list_children',
      {
        title: 'List OneDrive Folder Items',
        description: describeTool('List files and folders under a selected OneDrive folder.', [
          '查看OneDrive文件夹内容',
          '展开网盘目录',
          '列出文件夹里的文件'
        ]),
        inputSchema: {
          parentItemId: z.string().min(1).max(300).describe('Parent folder item ID.'),
          top: z.number().int().min(1).max(100).optional().describe('Number of children to return.')
        }
      },
      async ({ parentItemId, top }) => runTool(async () => listDriveChildren(config, parentItemId, top))
    );

    server.registerTool(
      'drive_list_recent',
      {
        title: 'List Recent OneDrive Items',
        description: describeTool('List files recently used by the signed-in user.', [
          '查看最近使用的文件',
          '列出最近打开的OneDrive文件',
          '我的最近文档'
        ]),
        inputSchema: {
          top: z.number().int().min(1).max(100).optional().describe('Number of recent items to return.')
        }
      },
      async ({ top }) => runTool(async () => listRecentDriveItems(config, top))
    );

    server.registerTool(
      'drive_search_items',
      {
        title: 'Search My OneDrive',
        description: describeTool(
          'Search files visible to the signed-in user in OneDrive. Requires delegated Files.Read.',
          ['搜索OneDrive文件', '查找网盘文件', '按文件名搜索']
        ),
        inputSchema: {
          query: z.string().min(1).max(200).describe('Search query, such as file name or document keyword.'),
          top: z.number().int().min(1).max(50).optional().describe('Number of search results to return, from 1 to 50.')
        }
      },
      async ({ query, top }) => runTool(async () => searchMyDrive(config, query, top))
    );

    server.registerTool(
      'drive_get_item',
      {
        title: 'Get OneDrive Item',
        description: describeTool(
          'Read metadata for a file or folder in the signed-in user OneDrive. Requires delegated Files.Read.',
          ['查看OneDrive文件详情', '查看网盘文件信息', '读取文件元数据']
        ),
        inputSchema: {
          itemId: z.string().min(1).max(300).describe('OneDrive item ID from a list or search result.')
        }
      },
      async ({ itemId }) => runTool(async () => getDriveItem(config, itemId))
    );

    server.registerTool(
      'drive_download_file',
      {
        title: 'Download OneDrive File',
        description: describeTool(
          'Download a OneDrive file and return its bytes as base64. The MCP response is limited to 10 MB.',
          ['下载OneDrive文件', '读取网盘文件内容', '获取文件base64']
        ),
        inputSchema: {
          itemId: z.string().min(1).max(300).describe('File item ID.')
        }
      },
      async ({ itemId }) => runTool(async () => downloadDriveFile(config, itemId))
    );

    server.registerTool(
      'drive_create_folder',
      {
        title: 'Create OneDrive Folder',
        description: describeTool(
          'Create a folder in the signed-in user OneDrive root or under a parent folder. Requires delegated Files.ReadWrite.',
          ['创建OneDrive文件夹', '新建网盘文件夹', '创建目录']
        ),
        inputSchema: {
          name: z.string().min(1).max(255).describe('New folder name.'),
          parentItemId: z
            .string()
            .min(1)
            .max(300)
            .optional()
            .describe('Optional parent folder ID. Defaults to the root folder.'),
          conflictBehavior: z
            .enum(['rename', 'replace', 'fail'])
            .optional()
            .describe('How to handle name conflicts. Defaults to rename.')
        }
      },
      async ({ name, parentItemId, conflictBehavior }) =>
        runTool(async () => createDriveFolder(config, name, parentItemId, conflictBehavior))
    );

    server.registerTool(
      'drive_upload_small_file',
      {
        title: 'Upload Small OneDrive File',
        description: describeTool(
          'Upload a small file to the signed-in user OneDrive. Use upload sessions for large files. Requires delegated Files.ReadWrite.',
          ['上传OneDrive文件', '上传网盘文件', '保存文件到OneDrive']
        ),
        inputSchema: {
          path: z
            .string()
            .min(1)
            .max(1000)
            .describe('Target path relative to the OneDrive root, for example Notes/demo.txt.'),
          content: z.string().min(1).describe('File content. Use raw text for text mode or base64 for base64 mode.'),
          contentEncoding: z.enum(['text', 'base64']).optional().describe('Content encoding. Defaults to text.'),
          contentType: z
            .string()
            .min(1)
            .max(200)
            .optional()
            .describe('Optional MIME type, for example text/plain or application/pdf.'),
          conflictBehavior: z
            .enum(['replace', 'rename', 'fail'])
            .optional()
            .describe('How to handle name conflicts. Defaults to replace.')
        }
      },
      async ({ path, content, contentEncoding, contentType, conflictBehavior }) =>
        runTool(async () => uploadSmallDriveFile(config, path, content, contentEncoding, contentType, conflictBehavior))
    );

    server.registerTool(
      'drive_rename_item',
      {
        title: 'Rename OneDrive Item',
        description: describeTool(
          'Rename a file or folder in the signed-in user OneDrive. Requires delegated Files.ReadWrite.',
          ['重命名OneDrive文件', '改文件名', '重命名文件夹']
        ),
        inputSchema: {
          itemId: z.string().min(1).max(300).describe('OneDrive item ID.'),
          newName: z.string().min(1).max(255).describe('New name.')
        }
      },
      async ({ itemId, newName }) => runTool(async () => renameDriveItem(config, itemId, newName))
    );

    server.registerTool(
      'drive_move_item',
      {
        title: 'Move OneDrive Item',
        description: describeTool(
          'Move a file or folder in the signed-in user OneDrive to a new parent folder. Requires delegated Files.ReadWrite.',
          ['移动OneDrive文件', '移动网盘文件', '把文件移到文件夹']
        ),
        inputSchema: {
          itemId: z.string().min(1).max(300).describe('OneDrive item ID to move.'),
          newParentItemId: z.string().min(1).max(300).describe('Destination parent folder ID.')
        }
      },
      async ({ itemId, newParentItemId }) => runTool(async () => moveDriveItem(config, itemId, newParentItemId))
    );

    server.registerTool(
      'drive_copy_item',
      {
        title: 'Copy OneDrive Item',
        description: describeTool(
          'Copy a OneDrive file or folder asynchronously to another folder. The response may include an operation monitor URL.',
          ['复制OneDrive文件', '复制网盘文件夹', '创建文件副本']
        ),
        inputSchema: {
          itemId: z.string().min(1).max(300).describe('Source item ID.'),
          parentItemId: z.string().min(1).max(300).describe('Destination parent folder ID.'),
          newName: z.string().min(1).max(255).optional().describe('Optional name for the copy.')
        }
      },
      async ({ itemId, parentItemId, newName }) =>
        runTool(async () => copyDriveItem(config, itemId, parentItemId, newName))
    );

    server.registerTool(
      'drive_delete_item',
      {
        title: 'Delete OneDrive Item',
        description: describeTool(
          'Delete a file or folder in the signed-in user OneDrive. Requires delegated Files.ReadWrite.',
          ['删除OneDrive文件', '删除网盘文件', '删除文件夹']
        ),
        inputSchema: {
          itemId: z.string().min(1).max(300).describe('OneDrive item ID.')
        }
      },
      async ({ itemId }) => runTool(async () => deleteDriveItem(config, itemId))
    );

    server.registerTool(
      'drive_create_share_link',
      {
        title: 'Create OneDrive Share Link',
        description: describeTool(
          'Create a sharing link for a file or folder in the signed-in user OneDrive. Requires delegated Files.ReadWrite. Actual sharing scope depends on tenant policy.',
          ['创建OneDrive分享链接', '分享网盘文件', '生成文件链接']
        ),
        inputSchema: {
          itemId: z.string().min(1).max(300).describe('OneDrive item ID.'),
          type: z.enum(['view', 'edit', 'embed']).optional().describe('Link type. Defaults to view.'),
          scope: z
            .enum(['anonymous', 'organization', 'users'])
            .optional()
            .describe('Sharing scope. Defaults to organization.')
        }
      },
      async ({ itemId, type, scope }) => runTool(async () => createDriveItemLink(config, itemId, type, scope))
    );

    server.registerTool(
      'drive_invite_item',
      {
        title: 'Invite People To OneDrive Item',
        description: describeTool(
          'Grant selected people read or write access to a OneDrive item and optionally send an invitation.',
          ['共享OneDrive文件给指定人员', '邀请同事访问文件', '授予文件权限']
        ),
        inputSchema: {
          itemId: z.string().min(1).max(300).describe('File or folder item ID.'),
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
      async ({ itemId, recipients, roles, message, requireSignIn, sendInvitation }) =>
        runTool(async () =>
          inviteDriveItem(config, itemId, recipients, roles, message, requireSignIn, sendInvitation)
        )
    );

    server.registerTool(
      'drive_list_permissions',
      {
        title: 'List OneDrive Item Permissions',
        description: describeTool(
          'List permission and sharing metadata for a file or folder in the signed-in user OneDrive. Requires delegated Files.Read.',
          ['查看OneDrive权限', '查看文件共享权限', '查看谁能访问文件']
        ),
        inputSchema: {
          itemId: z.string().min(1).max(300).describe('OneDrive item ID.')
        }
      },
      async ({ itemId }) => runTool(async () => listDriveItemPermissions(config, itemId))
    );

    server.registerTool(
      'drive_list_versions',
      {
        title: 'List OneDrive File Versions',
        description: describeTool('List retained versions of a OneDrive file.', [
          '查看OneDrive版本历史',
          '列出文件旧版本',
          '查看文档版本'
        ]),
        inputSchema: {
          itemId: z.string().min(1).max(300).describe('File item ID.')
        }
      },
      async ({ itemId }) => runTool(async () => listDriveVersions(config, itemId))
    );

    server.registerTool(
      'drive_restore_version',
      {
        title: 'Restore OneDrive File Version',
        description: describeTool('Restore a selected historical version of a OneDrive file.', [
          '恢复OneDrive文件版本',
          '回滚网盘文件',
          '还原文档旧版本'
        ]),
        inputSchema: {
          itemId: z.string().min(1).max(300).describe('File item ID.'),
          versionId: z.string().min(1).max(300).describe('Version ID from drive_list_versions.')
        }
      },
      async ({ itemId, versionId }) => runTool(async () => restoreDriveVersion(config, itemId, versionId))
    );
  }
};
