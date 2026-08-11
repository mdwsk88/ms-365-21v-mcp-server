import { z } from 'zod/v4';
import {
  getUserManager,
  listOrganizationUsers,
  listUserDirectReports,
  listUserMemberships
} from '../graph-extended.js';
import { getMe, getUser, searchUsers } from '../graph.js';
import { describeTool, runTool } from './results.js';
import type { ToolModule } from './types.js';

export const usersModule: ToolModule = {
  category: 'users',
  displayName: 'Users and Profiles',
  description: 'Signed-in user and organization profile tools.',
  requiredRole: 'mcp.users',
  toolNames: [
    'graph_get_me',
    'users_list',
    'users_search',
    'users_get_profile',
    'users_get_manager',
    'users_list_direct_reports',
    'users_list_memberships'
  ],
  register(server, config) {
    server.registerTool(
      'graph_get_me',
      {
        title: 'Get My Profile',
        description: describeTool(
          'Read the signed-in user profile from Microsoft Graph /me. Use for questions about account identity, email address, display name, job title, and department.',
          ['获取我的个人资料', '查看我的账号信息', '我是谁', '查看我的邮箱和部门']
        )
      },
      async () => runTool(async () => getMe(config))
    );

    server.registerTool(
      'users_list',
      {
        title: 'List Organization Users',
        description: describeTool(
          'List organization users with basic profile fields. Requires delegated User.ReadBasic.All.',
          ['列出公司用户', '查看组织通讯录', '浏览同事列表']
        ),
        inputSchema: {
          top: z.number().int().min(1).max(100).optional().describe('Number of users to return, from 1 to 100.')
        }
      },
      async ({ top }) => runTool(async () => listOrganizationUsers(config, top))
    );

    server.registerTool(
      'users_search',
      {
        title: 'Search Organization Users',
        description: describeTool(
          'Search tenant users by name, email, or UPN and return basic profile data. Requires delegated User.ReadBasic.All.',
          ['搜索公司用户', '查找同事', '按姓名查人', '按邮箱查人']
        ),
        inputSchema: {
          query: z.string().min(1).max(200).describe('Search query, such as display name, email prefix, or UPN.'),
          top: z.number().int().min(1).max(25).optional().describe('Number of users to return, from 1 to 25.')
        }
      },
      async ({ query, top }) => runTool(async () => searchUsers(config, query, top))
    );

    server.registerTool(
      'users_get_profile',
      {
        title: 'Get Organization User Profile',
        description: describeTool(
          'Read a tenant user profile, including job title, department, and contact fields. Requires delegated User.Read.All.',
          ['查看同事资料', '查看用户部门职位', '查看组织用户信息']
        ),
        inputSchema: {
          userIdOrPrincipalName: z.string().min(1).max(300).describe('User ID or userPrincipalName.')
        }
      },
      async ({ userIdOrPrincipalName }) => runTool(async () => getUser(config, userIdOrPrincipalName))
    );

    server.registerTool(
      'users_get_manager',
      {
        title: 'Get User Manager',
        description: describeTool(
          'Get the manager of the signed-in user or another organization user. Requires delegated User.Read.All.',
          ['查看我的经理', '查看同事的主管', '谁是这个用户的上级']
        ),
        inputSchema: {
          userIdOrPrincipalName: z
            .string()
            .min(1)
            .max(300)
            .optional()
            .describe('User ID or UPN. Omit to query the signed-in user.')
        }
      },
      async ({ userIdOrPrincipalName }) =>
        runTool(async () => getUserManager(config, userIdOrPrincipalName ?? 'me'))
    );

    server.registerTool(
      'users_list_direct_reports',
      {
        title: 'List User Direct Reports',
        description: describeTool(
          'List direct reports for the signed-in user or another organization user. Requires delegated User.Read.All.',
          ['查看我的下属', '列出直属团队成员', '查看某人的直接汇报人']
        ),
        inputSchema: {
          userIdOrPrincipalName: z
            .string()
            .min(1)
            .max(300)
            .optional()
            .describe('User ID or UPN. Omit to query the signed-in user.'),
          top: z.number().int().min(1).max(100).optional().describe('Number of direct reports to return.')
        }
      },
      async ({ userIdOrPrincipalName, top }) =>
        runTool(async () => listUserDirectReports(config, userIdOrPrincipalName ?? 'me', top))
    );

    server.registerTool(
      'users_list_memberships',
      {
        title: 'List User Memberships',
        description: describeTool(
          'List groups and directory objects that a user directly belongs to. Omit the user to query the signed-in user.',
          ['查看我加入的组', '查看用户所属组', '查询用户成员关系']
        ),
        inputSchema: {
          userIdOrPrincipalName: z
            .string()
            .min(1)
            .max(300)
            .optional()
            .describe('User ID or UPN. Omit to query the signed-in user.'),
          top: z.number().int().min(1).max(100).optional().describe('Number of memberships to return.')
        }
      },
      async ({ userIdOrPrincipalName, top }) =>
        runTool(async () => listUserMemberships(config, userIdOrPrincipalName ?? 'me', top))
    );
  }
};
