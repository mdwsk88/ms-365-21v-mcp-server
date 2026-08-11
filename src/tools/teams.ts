import { z } from 'zod/v4';
import {
  addChannelMember,
  addChatMember,
  addTeamMember,
  createTeamChannel,
  createTeamsChat,
  deleteTeamChannel,
  getTeam,
  getTeamChannel,
  listChannelMembers,
  listChatMembers,
  listTeamMembers,
  removeChannelMember,
  removeChatMember,
  removeTeamMember,
  softDeleteChannelMessage,
  softDeleteChatMessage,
  updateTeamChannel
} from '../graph-extended.js';
import {
  listChannelMessageReplies,
  listChannelMessages,
  listChatMessages,
  listMyChats,
  listMyJoinedTeams,
  listTeamChannels,
  replyToChannelMessage,
  sendChannelMessage,
  sendChatMessage
} from '../graph.js';
import { describeTool, runTool } from './results.js';
import type { ToolModule } from './types.js';

const conversationMemberSchema = z.object({
  userId: z.string().min(1).max(300).describe('Microsoft Entra user ID.'),
  role: z.enum(['owner', 'guest', 'member']).optional().describe('Membership role. Defaults to member.')
});

const chatMemberSchema = z.object({
  userId: z.string().min(1).max(300).describe('Microsoft Entra user ID.'),
  role: z.enum(['owner', 'guest']).optional().describe('Chat role. Defaults to owner; use guest for an in-tenant guest.')
});

export const teamsModule: ToolModule = {
  category: 'teams',
  displayName: 'Teams',
  description: 'Microsoft Teams channel and chat tools.',
  requiredRole: 'mcp.teams',
  toolNames: [
    'teams_list_joined_teams',
    'teams_get_team',
    'teams_list_team_members',
    'teams_add_team_member',
    'teams_remove_team_member',
    'teams_list_channels',
    'teams_get_channel',
    'teams_list_channel_members',
    'teams_create_channel',
    'teams_update_channel',
    'teams_delete_channel',
    'teams_add_channel_member',
    'teams_remove_channel_member',
    'teams_list_channel_messages',
    'teams_list_channel_message_replies',
    'teams_send_channel_message',
    'teams_reply_channel_message',
    'teams_delete_channel_message',
    'teams_list_chats',
    'teams_create_chat',
    'teams_list_chat_members',
    'teams_add_chat_member',
    'teams_remove_chat_member',
    'teams_list_chat_messages',
    'teams_send_chat_message',
    'teams_delete_chat_message'
  ],
  register(server, config) {
    server.registerTool(
      'teams_list_joined_teams',
      {
        title: 'List My Joined Teams',
        description: describeTool(
          'List Microsoft Teams teams that the signed-in user has directly joined. Requires delegated Team.ReadBasic.All.',
          ['查看我加入的Teams团队', '列出Teams团队', '查看团队列表']
        )
      },
      async () => runTool(async () => listMyJoinedTeams(config))
    );

    server.registerTool(
      'teams_get_team',
      {
        title: 'Get Teams Team',
        description: describeTool('Read settings and summary details for a Microsoft Teams team.', [
          '查看Teams团队详情',
          '读取团队设置',
          '查看团队信息'
        ]),
        inputSchema: {
          teamId: z.string().min(1).max(200).describe('Teams team ID.')
        }
      },
      async ({ teamId }) => runTool(async () => getTeam(config, teamId))
    );

    server.registerTool(
      'teams_list_team_members',
      {
        title: 'List Teams Team Members',
        description: describeTool('List members and owners of a Microsoft Teams team.', [
          '查看Teams团队成员',
          '列出团队所有者',
          '谁在这个团队里'
        ]),
        inputSchema: {
          teamId: z.string().min(1).max(200).describe('Teams team ID.'),
          top: z.number().int().min(1).max(999).optional().describe('Number of members to return.')
        }
      },
      async ({ teamId, top }) => runTool(async () => listTeamMembers(config, teamId, top))
    );

    server.registerTool(
      'teams_add_team_member',
      {
        title: 'Add Teams Team Member',
        description: describeTool(
          'Add a Microsoft Entra user as a team member, owner, or guest. Requires TeamMember.ReadWrite.All and the caller must have authority in the team.',
          ['添加Teams团队成员', '给团队加所有者', '邀请用户加入团队']
        ),
        inputSchema: {
          teamId: z.string().min(1).max(200).describe('Teams team ID.'),
          member: conversationMemberSchema.describe('User and role to add.')
        }
      },
      async ({ teamId, member }) => runTool(async () => addTeamMember(config, teamId, member))
    );

    server.registerTool(
      'teams_remove_team_member',
      {
        title: 'Remove Teams Team Member',
        description: describeTool('Remove a member from a Microsoft Teams team.', [
          '移除Teams团队成员',
          '把用户移出团队',
          '删除团队成员'
        ]),
        inputSchema: {
          teamId: z.string().min(1).max(200).describe('Teams team ID.'),
          membershipId: z.string().min(1).max(1000).describe('Opaque membership ID from teams_list_team_members.')
        }
      },
      async ({ teamId, membershipId }) => runTool(async () => removeTeamMember(config, teamId, membershipId))
    );

    server.registerTool(
      'teams_list_channels',
      {
        title: 'List Team Channels',
        description: describeTool(
          'List channels in a Microsoft Teams team. Use teams_list_joined_teams first to obtain a teamId. Requires delegated Channel.ReadBasic.All.',
          ['查看Teams频道', '列出团队频道', '查看团队有哪些频道']
        ),
        inputSchema: {
          teamId: z.string().min(1).max(200).describe('Teams team ID, usually from teams_list_joined_teams.')
        }
      },
      async ({ teamId }) => runTool(async () => listTeamChannels(config, teamId))
    );

    server.registerTool(
      'teams_get_channel',
      {
        title: 'Get Teams Channel',
        description: describeTool('Read details for one Microsoft Teams channel.', [
          '查看Teams频道详情',
          '读取频道信息',
          '查看频道类型'
        ]),
        inputSchema: {
          teamId: z.string().min(1).max(200).describe('Teams team ID.'),
          channelId: z.string().min(1).max(300).describe('Teams channel ID.')
        }
      },
      async ({ teamId, channelId }) => runTool(async () => getTeamChannel(config, teamId, channelId))
    );

    server.registerTool(
      'teams_list_channel_members',
      {
        title: 'List Teams Channel Members',
        description: describeTool('List direct members of a standard, private, or shared Teams channel.', [
          '查看Teams频道成员',
          '列出私有频道用户',
          '谁在这个频道里'
        ]),
        inputSchema: {
          teamId: z.string().min(1).max(200).describe('Teams team ID.'),
          channelId: z.string().min(1).max(300).describe('Teams channel ID.'),
          top: z.number().int().min(1).max(999).optional().describe('Number of members to return.')
        }
      },
      async ({ teamId, channelId, top }) =>
        runTool(async () => listChannelMembers(config, teamId, channelId, top))
    );

    server.registerTool(
      'teams_create_channel',
      {
        title: 'Create Teams Channel',
        description: describeTool(
          'Create a standard or private Teams channel. Shared-channel creation is intentionally excluded because Microsoft does not support it in 21V.',
          ['创建Teams频道', '新建私有频道', '给团队添加频道']
        ),
        inputSchema: {
          teamId: z.string().min(1).max(200).describe('Teams team ID.'),
          displayName: z.string().min(1).max(50).describe('Channel display name.'),
          description: z.string().max(1000).optional().describe('Optional channel description.'),
          membershipType: z.enum(['standard', 'private']).optional().describe('Channel type. Defaults to standard.'),
          members: z
            .array(conversationMemberSchema)
            .min(1)
            .max(200)
            .optional()
            .describe('Initial private-channel members. Include an owner for a private channel.')
        }
      },
      async ({ teamId, displayName, description, membershipType, members }) =>
        runTool(async () =>
          createTeamChannel(config, teamId, displayName, description, membershipType, members)
        )
    );

    server.registerTool(
      'teams_update_channel',
      {
        title: 'Update Teams Channel',
        description: describeTool('Update the name or description of a Microsoft Teams channel.', [
          '修改Teams频道',
          '重命名频道',
          '更新频道描述'
        ]),
        inputSchema: {
          teamId: z.string().min(1).max(200).describe('Teams team ID.'),
          channelId: z.string().min(1).max(300).describe('Teams channel ID.'),
          displayName: z.string().min(1).max(50).optional().describe('New channel name.'),
          description: z.string().max(1000).optional().describe('New description.')
        }
      },
      async ({ teamId, channelId, displayName, description }) =>
        runTool(async () => updateTeamChannel(config, teamId, channelId, { displayName, description }))
    );

    server.registerTool(
      'teams_delete_channel',
      {
        title: 'Delete Teams Channel',
        description: describeTool('Delete a Microsoft Teams channel.', ['删除Teams频道', '移除团队频道']),
        inputSchema: {
          teamId: z.string().min(1).max(200).describe('Teams team ID.'),
          channelId: z.string().min(1).max(300).describe('Channel ID to delete.')
        }
      },
      async ({ teamId, channelId }) => runTool(async () => deleteTeamChannel(config, teamId, channelId))
    );

    server.registerTool(
      'teams_add_channel_member',
      {
        title: 'Add Teams Channel Member',
        description: describeTool('Add a member or owner to a private Teams channel.', [
          '添加Teams频道成员',
          '给私有频道加用户',
          '添加频道所有者'
        ]),
        inputSchema: {
          teamId: z.string().min(1).max(200).describe('Teams team ID.'),
          channelId: z.string().min(1).max(300).describe('Teams channel ID.'),
          member: conversationMemberSchema.describe('User and role to add.')
        }
      },
      async ({ teamId, channelId, member }) =>
        runTool(async () => addChannelMember(config, teamId, channelId, member))
    );

    server.registerTool(
      'teams_remove_channel_member',
      {
        title: 'Remove Teams Channel Member',
        description: describeTool('Remove a direct member from a private Teams channel.', [
          '移除Teams频道成员',
          '把用户移出私有频道',
          '删除频道成员'
        ]),
        inputSchema: {
          teamId: z.string().min(1).max(200).describe('Teams team ID.'),
          channelId: z.string().min(1).max(300).describe('Teams channel ID.'),
          membershipId: z.string().min(1).max(1000).describe('Opaque membership ID.')
        }
      },
      async ({ teamId, channelId, membershipId }) =>
        runTool(async () => removeChannelMember(config, teamId, channelId, membershipId))
    );

    server.registerTool(
      'teams_list_channel_messages',
      {
        title: 'List Teams Channel Messages',
        description: describeTool(
          'List recent messages in a Teams channel. Requires delegated ChannelMessage.Read.All and usually tenant admin consent.',
          ['查看Teams频道消息', '读取频道聊天记录', '查看团队频道最近消息']
        ),
        inputSchema: {
          teamId: z.string().min(1).max(200).describe('Teams team ID.'),
          channelId: z.string().min(1).max(300).describe('Teams channel ID, usually from teams_list_channels.'),
          top: z.number().int().min(1).max(50).optional().describe('Number of messages to return, from 1 to 50.')
        }
      },
      async ({ teamId, channelId, top }) => runTool(async () => listChannelMessages(config, teamId, channelId, top))
    );

    server.registerTool(
      'teams_list_channel_message_replies',
      {
        title: 'List Teams Channel Message Replies',
        description: describeTool(
          'List replies under a Teams channel message. Requires delegated ChannelMessage.Read.All and usually tenant admin consent.',
          ['查看Teams频道消息回复', '查看线程回复', '读取频道消息评论']
        ),
        inputSchema: {
          teamId: z.string().min(1).max(200).describe('Teams team ID.'),
          channelId: z.string().min(1).max(300).describe('Teams channel ID.'),
          messageId: z.string().min(1).max(300).describe('Channel message ID.'),
          top: z.number().int().min(1).max(50).optional().describe('Number of replies to return, from 1 to 50.')
        }
      },
      async ({ teamId, channelId, messageId, top }) =>
        runTool(async () => listChannelMessageReplies(config, teamId, channelId, messageId, top))
    );

    server.registerTool(
      'teams_send_channel_message',
      {
        title: 'Send Teams Channel Message',
        description: describeTool(
          'Send a message to a Teams channel as the signed-in user. Requires delegated ChannelMessage.Send.',
          ['发送Teams频道消息', '在频道发消息', '给团队频道发通知']
        ),
        inputSchema: {
          teamId: z.string().min(1).max(200).describe('Teams team ID.'),
          channelId: z.string().min(1).max(300).describe('Teams channel ID.'),
          content: z.string().min(1).max(20000).describe('Message content.'),
          contentIsHtml: z.boolean().optional().describe('Whether the content is HTML. Defaults to false.')
        }
      },
      async ({ teamId, channelId, content, contentIsHtml }) =>
        runTool(async () => sendChannelMessage(config, teamId, channelId, content, contentIsHtml))
    );

    server.registerTool(
      'teams_reply_channel_message',
      {
        title: 'Reply To Teams Channel Message',
        description: describeTool(
          'Reply to a Teams channel message as the signed-in user. Requires delegated ChannelMessage.Send.',
          ['回复Teams频道消息', '回复频道线程', '在频道消息下回复']
        ),
        inputSchema: {
          teamId: z.string().min(1).max(200).describe('Teams team ID.'),
          channelId: z.string().min(1).max(300).describe('Teams channel ID.'),
          messageId: z.string().min(1).max(300).describe('Channel message ID.'),
          content: z.string().min(1).max(20000).describe('Reply content.'),
          contentIsHtml: z.boolean().optional().describe('Whether the content is HTML. Defaults to false.')
        }
      },
      async ({ teamId, channelId, messageId, content, contentIsHtml }) =>
        runTool(async () => replyToChannelMessage(config, teamId, channelId, messageId, content, contentIsHtml))
    );

    server.registerTool(
      'teams_delete_channel_message',
      {
        title: 'Delete Teams Channel Message',
        description: describeTool(
          'Soft-delete a Teams channel message or reply. Microsoft Graph and tenant policy enforce who may delete it.',
          ['删除Teams频道消息', '撤回频道消息', '删除频道回复']
        ),
        inputSchema: {
          teamId: z.string().min(1).max(200).describe('Teams team ID.'),
          channelId: z.string().min(1).max(300).describe('Teams channel ID.'),
          messageId: z.string().min(1).max(300).describe('Root channel message ID.'),
          replyId: z.string().min(1).max(300).optional().describe('Optional reply ID. Omit to delete the root message.')
        }
      },
      async ({ teamId, channelId, messageId, replyId }) =>
        runTool(async () => softDeleteChannelMessage(config, teamId, channelId, messageId, replyId))
    );

    server.registerTool(
      'teams_list_chats',
      {
        title: 'List My Teams Chats',
        description: describeTool(
          'List one-on-one and group Teams chats involving the signed-in user. Requires delegated Chat.Read.',
          ['查看Teams聊天', '列出Teams会话', '查看群聊和私聊']
        ),
        inputSchema: {
          top: z.number().int().min(1).max(50).optional().describe('Number of chats to return, from 1 to 50.')
        }
      },
      async ({ top }) => runTool(async () => listMyChats(config, top))
    );

    server.registerTool(
      'teams_create_chat',
      {
        title: 'Create Teams Chat',
        description: describeTool(
          'Create a one-on-one or group Teams chat. Include every participant, including the signed-in user.',
          ['创建Teams聊天', '发起群聊', '创建一对一聊天']
        ),
        inputSchema: {
          chatType: z.enum(['oneOnOne', 'group']).describe('Chat type.'),
          members: z.array(chatMemberSchema).min(2).max(250).describe('All chat participants.'),
          topic: z.string().min(1).max(250).optional().describe('Group-chat topic. Ignored for one-on-one chat.')
        }
      },
      async ({ chatType, members, topic }) => runTool(async () => createTeamsChat(config, chatType, members, topic))
    );

    server.registerTool(
      'teams_list_chat_members',
      {
        title: 'List Teams Chat Members',
        description: describeTool('List all members of a Teams chat.', ['查看Teams聊天成员', '列出群聊人员', '谁在这个聊天里']),
        inputSchema: {
          chatId: z.string().min(1).max(300).describe('Teams chat ID.')
        }
      },
      async ({ chatId }) => runTool(async () => listChatMembers(config, chatId))
    );

    server.registerTool(
      'teams_add_chat_member',
      {
        title: 'Add Teams Chat Member',
        description: describeTool('Add a user to an existing Teams group chat.', [
          '添加Teams群聊成员',
          '邀请用户加入聊天',
          '给群聊加人'
        ]),
        inputSchema: {
          chatId: z.string().min(1).max(300).describe('Teams chat ID.'),
          member: chatMemberSchema.describe('User to add.'),
          visibleHistoryStartDateTime: z
            .string()
            .min(1)
            .max(80)
            .optional()
            .describe('Optional ISO date-time controlling visible chat history.')
        }
      },
      async ({ chatId, member, visibleHistoryStartDateTime }) =>
        runTool(async () => addChatMember(config, chatId, member, visibleHistoryStartDateTime))
    );

    server.registerTool(
      'teams_remove_chat_member',
      {
        title: 'Remove Teams Chat Member',
        description: describeTool('Remove a member from a Teams group chat.', [
          '移除Teams群聊成员',
          '把用户移出聊天',
          '删除群聊成员'
        ]),
        inputSchema: {
          chatId: z.string().min(1).max(300).describe('Teams chat ID.'),
          membershipId: z.string().min(1).max(1000).describe('Opaque membership ID from teams_list_chat_members.')
        }
      },
      async ({ chatId, membershipId }) => runTool(async () => removeChatMember(config, chatId, membershipId))
    );

    server.registerTool(
      'teams_list_chat_messages',
      {
        title: 'List Teams Chat Messages',
        description: describeTool('List recent messages in a Teams chat. Requires delegated Chat.Read.', [
          '查看Teams聊天消息',
          '读取聊天记录',
          '查看群聊消息'
        ]),
        inputSchema: {
          chatId: z.string().min(1).max(300).describe('Chat ID from teams_list_chats.'),
          top: z.number().int().min(1).max(50).optional().describe('Number of messages to return, from 1 to 50.')
        }
      },
      async ({ chatId, top }) => runTool(async () => listChatMessages(config, chatId, top))
    );

    server.registerTool(
      'teams_send_chat_message',
      {
        title: 'Send Teams Chat Message',
        description: describeTool(
          'Send a message to a Teams chat as the signed-in user. Requires delegated ChatMessage.Send.',
          ['发送Teams聊天消息', '给Teams群聊发消息', '发送私聊消息']
        ),
        inputSchema: {
          chatId: z.string().min(1).max(300).describe('Chat ID.'),
          content: z.string().min(1).max(20000).describe('Message content.'),
          contentIsHtml: z.boolean().optional().describe('Whether the content is HTML. Defaults to false.')
        }
      },
      async ({ chatId, content, contentIsHtml }) =>
        runTool(async () => sendChatMessage(config, chatId, content, contentIsHtml))
    );

    server.registerTool(
      'teams_delete_chat_message',
      {
        title: 'Delete Teams Chat Message',
        description: describeTool(
          'Soft-delete a Teams chat message. Microsoft Graph and tenant policy enforce who may delete it.',
          ['删除Teams聊天消息', '撤回私聊消息', '删除群聊消息']
        ),
        inputSchema: {
          chatId: z.string().min(1).max(300).describe('Teams chat ID.'),
          messageId: z.string().min(1).max(300).describe('Chat message ID.')
        }
      },
      async ({ chatId, messageId }) => runTool(async () => softDeleteChatMessage(config, chatId, messageId))
    );
  }
};
