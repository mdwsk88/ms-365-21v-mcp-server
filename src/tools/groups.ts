import { z } from 'zod/v4';
import {
  checkMyGroupMemberships,
  getGroup,
  listGroupMembers,
  listGroupOwners,
  listGroups
} from '../graph-extended.js';
import { describeTool, runTool } from './results.js';
import type { ToolModule } from './types.js';

export const groupsModule: ToolModule = {
  category: 'groups',
  displayName: 'Groups',
  description: 'Microsoft Entra and Microsoft 365 group discovery tools.',
  requiredRole: 'mcp.groups',
  toolNames: [
    'groups_list',
    'groups_get',
    'groups_list_members',
    'groups_list_owners',
    'groups_check_my_memberships'
  ],
  register(server, config) {
    server.registerTool(
      'groups_list',
      {
        title: 'List Organization Groups',
        description: describeTool(
          'List Microsoft Entra security groups and Microsoft 365 groups. Optionally filter by display-name prefix. Requires delegated GroupMember.Read.All.',
          ['列出公司组', '搜索安全组', '查看Microsoft 365组']
        ),
        inputSchema: {
          query: z.string().min(1).max(200).optional().describe('Optional display-name prefix.'),
          top: z.number().int().min(1).max(100).optional().describe('Number of groups to return.')
        }
      },
      async ({ query, top }) => runTool(async () => listGroups(config, query, top))
    );

    server.registerTool(
      'groups_get',
      {
        title: 'Get Organization Group',
        description: describeTool('Read details for a Microsoft Entra or Microsoft 365 group.', [
          '查看组详情',
          '查看安全组信息',
          '查看Microsoft 365组信息'
        ]),
        inputSchema: {
          groupId: z.string().min(1).max(300).describe('Group ID.')
        }
      },
      async ({ groupId }) => runTool(async () => getGroup(config, groupId))
    );

    server.registerTool(
      'groups_list_members',
      {
        title: 'List Group Members',
        description: describeTool(
          'List direct members of a group. Hidden-membership groups may require additional tenant permissions.',
          ['查看组成员', '列出安全组用户', '谁在这个组里']
        ),
        inputSchema: {
          groupId: z.string().min(1).max(300).describe('Group ID.'),
          top: z.number().int().min(1).max(100).optional().describe('Number of members to return.')
        }
      },
      async ({ groupId, top }) => runTool(async () => listGroupMembers(config, groupId, top))
    );

    server.registerTool(
      'groups_list_owners',
      {
        title: 'List Group Owners',
        description: describeTool('List owners of a Microsoft Entra or Microsoft 365 group.', [
          '查看组所有者',
          '查看组管理员',
          '谁负责这个组'
        ]),
        inputSchema: {
          groupId: z.string().min(1).max(300).describe('Group ID.'),
          top: z.number().int().min(1).max(100).optional().describe('Number of owners to return.')
        }
      },
      async ({ groupId, top }) => runTool(async () => listGroupOwners(config, groupId, top))
    );

    server.registerTool(
      'groups_check_my_memberships',
      {
        title: 'Check My Group Memberships',
        description: describeTool(
          'Check which of up to 20 supplied group IDs contain the signed-in user, including transitive membership.',
          ['检查我是否在这些组', '验证我的组成员身份', '检查用户权限组']
        ),
        inputSchema: {
          groupIds: z.array(z.string().min(1).max(300)).min(1).max(20).describe('Group IDs to check.')
        }
      },
      async ({ groupIds }) => runTool(async () => checkMyGroupMemberships(config, groupIds))
    );
  }
};
