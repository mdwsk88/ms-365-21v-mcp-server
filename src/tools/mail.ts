import { z } from 'zod/v4';
import {
  addMessageFileAttachment,
  copyMessage,
  createMailFolder,
  deleteMailFolder,
  deleteMessageAttachment,
  getMailFolder,
  getMessageAttachment,
  listChildMailFolders,
  listFolderMessages,
  updateMailFolder
} from '../graph-extended.js';
import {
  createDraftMessage,
  deleteMessage,
  forwardMessage,
  getMessage,
  listMailFolders,
  listMessageAttachments,
  listMyMessages,
  moveMessage,
  replyToMessage,
  searchMyMessages,
  sendDraftMessage,
  sendMail,
  updateMessageReadState
} from '../graph.js';
import { recipientSchema } from './schemas.js';
import { describeTool, runTool } from './results.js';
import type { ToolModule } from './types.js';

export const mailModule: ToolModule = {
  category: 'mail',
  displayName: 'Mail',
  description: 'Outlook mailbox read and write tools.',
  requiredRole: 'mcp.mail',
  toolNames: [
    'mail_list_messages',
    'mail_search_messages',
    'mail_get_message',
    'mail_list_folders',
    'mail_get_folder',
    'mail_list_child_folders',
    'mail_list_folder_messages',
    'mail_create_folder',
    'mail_update_folder',
    'mail_delete_folder',
    'mail_list_attachments',
    'mail_get_attachment',
    'mail_add_file_attachment',
    'mail_delete_attachment',
    'mail_send',
    'mail_create_draft',
    'mail_send_draft',
    'mail_reply',
    'mail_forward',
    'mail_set_read_state',
    'mail_move_message',
    'mail_copy_message',
    'mail_delete_message'
  ],
  register(server, config) {
    server.registerTool(
      'mail_list_messages',
      {
        title: 'List My Mail',
        description: describeTool(
          'List recent messages in the signed-in user mailbox. Supports recent mail and unread mail queries. Requires delegated Mail.Read.',
          ['查看我的邮件', '看最近邮件', '查看未读邮件', '列出收件箱邮件']
        ),
        inputSchema: {
          top: z.number().int().min(1).max(50).optional().describe('Number of messages to return, from 1 to 50.'),
          unreadOnly: z.boolean().optional().describe('Whether to return unread messages only.')
        }
      },
      async ({ top, unreadOnly }) => runTool(async () => listMyMessages(config, top, unreadOnly))
    );

    server.registerTool(
      'mail_search_messages',
      {
        title: 'Search My Mail',
        description: describeTool(
          'Search messages in the signed-in user mailbox by keyword. Requires delegated Mail.Read.',
          ['搜索邮件', '按关键词查邮件', '查找某封邮件']
        ),
        inputSchema: {
          query: z.string().min(1).max(200).describe('Search query, such as sender, subject, or body keywords.'),
          top: z.number().int().min(1).max(50).optional().describe('Number of messages to return, from 1 to 50.')
        }
      },
      async ({ query, top }) => runTool(async () => searchMyMessages(config, query, top))
    );

    server.registerTool(
      'mail_get_message',
      {
        title: 'Get Mail Details',
        description: describeTool(
          'Read detailed content for a message, including body, sender, recipients, and attachment flag. Requires delegated Mail.Read.',
          ['查看邮件详情', '读取邮件正文', '查看邮件内容', '查看邮件收件人']
        ),
        inputSchema: {
          messageId: z.string().min(1).max(300).describe('Message ID from a mail list or search result.')
        }
      },
      async ({ messageId }) => runTool(async () => getMessage(config, messageId))
    );

    server.registerTool(
      'mail_list_folders',
      {
        title: 'List Mail Folders',
        description: describeTool(
          'List mailbox folders for the signed-in user. Use this to choose a destination when moving messages. Requires delegated Mail.Read.',
          ['查看邮件文件夹', '列出邮箱文件夹', '查看收件箱草稿箱已发送']
        ),
        inputSchema: {
          top: z.number().int().min(1).max(100).optional().describe('Number of folders to return, from 1 to 100.')
        }
      },
      async ({ top }) => runTool(async () => listMailFolders(config, top))
    );

    server.registerTool(
      'mail_get_folder',
      {
        title: 'Get Mail Folder',
        description: describeTool('Read mailbox-folder details and item counts. Requires delegated Mail.Read.', [
          '查看邮件文件夹详情',
          '查看文件夹未读数量',
          '读取邮箱目录信息'
        ]),
        inputSchema: {
          folderId: z.string().min(1).max(300).describe('Folder ID or well-known name such as inbox or drafts.')
        }
      },
      async ({ folderId }) => runTool(async () => getMailFolder(config, folderId))
    );

    server.registerTool(
      'mail_list_child_folders',
      {
        title: 'List Child Mail Folders',
        description: describeTool('List child folders under a mailbox folder. Requires delegated Mail.Read.', [
          '查看子邮件文件夹',
          '展开邮箱目录',
          '列出文件夹下的目录'
        ]),
        inputSchema: {
          folderId: z.string().min(1).max(300).describe('Parent folder ID or well-known folder name.'),
          top: z.number().int().min(1).max(100).optional().describe('Number of child folders to return.'),
          includeHidden: z.boolean().optional().describe('Whether hidden folders should be included.')
        }
      },
      async ({ folderId, top, includeHidden }) =>
        runTool(async () => listChildMailFolders(config, folderId, top, includeHidden))
    );

    server.registerTool(
      'mail_list_folder_messages',
      {
        title: 'List Mail In Folder',
        description: describeTool('List recent messages in a selected mailbox folder. Requires delegated Mail.Read.', [
          '查看指定文件夹邮件',
          '查看归档邮件',
          '列出草稿箱邮件'
        ]),
        inputSchema: {
          folderId: z.string().min(1).max(300).describe('Folder ID or well-known folder name.'),
          top: z.number().int().min(1).max(50).optional().describe('Number of messages to return.'),
          unreadOnly: z.boolean().optional().describe('Whether to return unread messages only.')
        }
      },
      async ({ folderId, top, unreadOnly }) =>
        runTool(async () => listFolderMessages(config, folderId, top, unreadOnly))
    );

    server.registerTool(
      'mail_create_folder',
      {
        title: 'Create Mail Folder',
        description: describeTool('Create a top-level or child mailbox folder. Requires delegated Mail.ReadWrite.', [
          '创建邮件文件夹',
          '新建邮箱目录',
          '创建子文件夹'
        ]),
        inputSchema: {
          displayName: z.string().min(1).max(255).describe('Folder display name.'),
          parentFolderId: z.string().min(1).max(300).optional().describe('Optional parent folder ID.'),
          isHidden: z.boolean().optional().describe('Whether the folder is hidden. This cannot be changed later.')
        }
      },
      async ({ displayName, parentFolderId, isHidden }) =>
        runTool(async () => createMailFolder(config, displayName, parentFolderId, isHidden))
    );

    server.registerTool(
      'mail_update_folder',
      {
        title: 'Rename Mail Folder',
        description: describeTool('Rename a mailbox folder. Requires delegated Mail.ReadWrite.', [
          '重命名邮件文件夹',
          '修改邮箱目录名称'
        ]),
        inputSchema: {
          folderId: z.string().min(1).max(300).describe('Folder ID.'),
          displayName: z.string().min(1).max(255).describe('New folder display name.')
        }
      },
      async ({ folderId, displayName }) => runTool(async () => updateMailFolder(config, folderId, displayName))
    );

    server.registerTool(
      'mail_delete_folder',
      {
        title: 'Delete Mail Folder',
        description: describeTool('Delete a mailbox folder and its contents. Requires delegated Mail.ReadWrite.', [
          '删除邮件文件夹',
          '移除邮箱目录'
        ]),
        inputSchema: {
          folderId: z.string().min(1).max(300).describe('Folder ID to delete.')
        }
      },
      async ({ folderId }) => runTool(async () => deleteMailFolder(config, folderId))
    );

    server.registerTool(
      'mail_list_attachments',
      {
        title: 'List Mail Attachments',
        description: describeTool('List attachment metadata for a message. Requires delegated Mail.Read.', [
          '查看邮件附件',
          '列出邮件附件',
          '邮件有没有附件'
        ]),
        inputSchema: {
          messageId: z.string().min(1).max(300).describe('Message ID.')
        }
      },
      async ({ messageId }) => runTool(async () => listMessageAttachments(config, messageId))
    );

    server.registerTool(
      'mail_get_attachment',
      {
        title: 'Get Mail Attachment',
        description: describeTool(
          'Read one mail attachment. Set raw=true to return file bytes as base64, limited to 10 MB. Requires delegated Mail.Read.',
          ['读取邮件附件', '下载邮件附件', '查看附件内容']
        ),
        inputSchema: {
          messageId: z.string().min(1).max(300).describe('Message ID.'),
          attachmentId: z.string().min(1).max(500).describe('Attachment ID.'),
          raw: z.boolean().optional().describe('Return raw attachment bytes as base64. Defaults to false.')
        }
      },
      async ({ messageId, attachmentId, raw }) =>
        runTool(async () => getMessageAttachment(config, messageId, attachmentId, raw))
    );

    server.registerTool(
      'mail_add_file_attachment',
      {
        title: 'Add File Attachment To Mail',
        description: describeTool(
          'Add a base64-encoded file under 3 MB to an existing draft message. Requires delegated Mail.ReadWrite.',
          ['给邮件添加附件', '给草稿添加文件', '附加文件到邮件']
        ),
        inputSchema: {
          messageId: z.string().min(1).max(300).describe('Draft message ID.'),
          name: z.string().min(1).max(255).describe('Attachment file name.'),
          contentType: z.string().min(1).max(200).describe('MIME type.'),
          contentBase64: z.string().min(1).max(4_200_000).describe('Base64-encoded content, under 3 MB decoded.'),
          isInline: z.boolean().optional().describe('Whether this is an inline attachment.'),
          contentId: z.string().min(1).max(255).optional().describe('Optional inline content ID.')
        }
      },
      async ({ messageId, name, contentType, contentBase64, isInline, contentId }) =>
        runTool(async () =>
          addMessageFileAttachment(config, messageId, name, contentType, contentBase64, isInline, contentId)
        )
    );

    server.registerTool(
      'mail_delete_attachment',
      {
        title: 'Delete Mail Attachment',
        description: describeTool('Delete an attachment from a draft message. Requires delegated Mail.ReadWrite.', [
          '删除邮件附件',
          '从草稿移除附件'
        ]),
        inputSchema: {
          messageId: z.string().min(1).max(300).describe('Draft message ID.'),
          attachmentId: z.string().min(1).max(500).describe('Attachment ID to delete.')
        }
      },
      async ({ messageId, attachmentId }) =>
        runTool(async () => deleteMessageAttachment(config, messageId, attachmentId))
    );

    server.registerTool(
      'mail_send',
      {
        title: 'Send Mail',
        description: describeTool('Send an email as the signed-in user. Requires delegated Mail.Send.', [
          '发送邮件',
          '帮我发邮件',
          '给某人发邮件'
        ]),
        inputSchema: {
          subject: z.string().min(1).max(255).describe('Email subject.'),
          body: z.string().min(1).max(20000).describe('Email body.'),
          to: z.array(recipientSchema).min(1).max(50).describe('To recipients.'),
          cc: z.array(recipientSchema).max(50).optional().describe('Cc recipients.'),
          bcc: z.array(recipientSchema).max(50).optional().describe('Bcc recipients.'),
          bodyIsHtml: z.boolean().optional().describe('Whether the body is HTML. Defaults to false.'),
          saveToSentItems: z
            .boolean()
            .optional()
            .describe('Whether to save the message to Sent Items. Defaults to true.')
        }
      },
      async (input) => runTool(async () => sendMail(config, input))
    );

    server.registerTool(
      'mail_create_draft',
      {
        title: 'Create Mail Draft',
        description: describeTool('Create an email draft without sending it. Requires delegated Mail.ReadWrite.', [
          '创建邮件草稿',
          '写邮件草稿',
          '先不要发送邮件'
        ]),
        inputSchema: {
          subject: z.string().min(1).max(255).describe('Email subject.'),
          body: z.string().min(1).max(20000).describe('Email body.'),
          to: z.array(recipientSchema).max(50).optional().describe('To recipients.'),
          cc: z.array(recipientSchema).max(50).optional().describe('Cc recipients.'),
          bcc: z.array(recipientSchema).max(50).optional().describe('Bcc recipients.'),
          bodyIsHtml: z.boolean().optional().describe('Whether the body is HTML. Defaults to false.')
        }
      },
      async (input) => runTool(async () => createDraftMessage(config, input))
    );

    server.registerTool(
      'mail_send_draft',
      {
        title: 'Send Mail Draft',
        description: describeTool('Send an existing draft message. Requires delegated Mail.Send.', [
          '发送草稿',
          '发送邮件草稿',
          '把草稿发出去'
        ]),
        inputSchema: {
          messageId: z.string().min(1).max(300).describe('Draft message ID.')
        }
      },
      async ({ messageId }) => runTool(async () => sendDraftMessage(config, messageId))
    );

    server.registerTool(
      'mail_reply',
      {
        title: 'Reply To Mail',
        description: describeTool('Reply to an existing message. Requires delegated Mail.Send.', [
          '回复邮件',
          '回这封邮件',
          '给邮件写回复'
        ]),
        inputSchema: {
          messageId: z.string().min(1).max(300).describe('Message ID.'),
          comment: z.string().min(1).max(10000).describe('Reply comment.')
        }
      },
      async ({ messageId, comment }) => runTool(async () => replyToMessage(config, messageId, comment))
    );

    server.registerTool(
      'mail_forward',
      {
        title: 'Forward Mail',
        description: describeTool(
          'Forward an existing message to one or more recipients. Requires delegated Mail.Send.',
          ['转发邮件', '把邮件转给别人', '转发这封邮件']
        ),
        inputSchema: {
          messageId: z.string().min(1).max(300).describe('Message ID.'),
          comment: z.string().max(10000).describe('Optional forwarding comment. May be an empty string.'),
          to: z.array(recipientSchema).min(1).max(50).describe('Forward recipients.')
        }
      },
      async ({ messageId, comment, to }) => runTool(async () => forwardMessage(config, messageId, comment, to))
    );

    server.registerTool(
      'mail_set_read_state',
      {
        title: 'Set Mail Read State',
        description: describeTool('Mark a message as read or unread. Requires delegated Mail.ReadWrite.', [
          '标记邮件已读',
          '标记邮件未读',
          '更改邮件读取状态'
        ]),
        inputSchema: {
          messageId: z.string().min(1).max(300).describe('Message ID.'),
          isRead: z.boolean().describe('true marks the message as read; false marks it as unread.')
        }
      },
      async ({ messageId, isRead }) => runTool(async () => updateMessageReadState(config, messageId, isRead))
    );

    server.registerTool(
      'mail_move_message',
      {
        title: 'Move Mail',
        description: describeTool('Move a message to a target mailbox folder. Requires delegated Mail.ReadWrite.', [
          '移动邮件',
          '把邮件移到文件夹',
          '整理邮件'
        ]),
        inputSchema: {
          messageId: z.string().min(1).max(300).describe('Message ID.'),
          destinationFolderId: z.string().min(1).max(300).describe('Destination folder ID from mail_list_folders.')
        }
      },
      async ({ messageId, destinationFolderId }) =>
        runTool(async () => moveMessage(config, messageId, destinationFolderId))
    );

    server.registerTool(
      'mail_copy_message',
      {
        title: 'Copy Mail',
        description: describeTool('Copy a message into another mailbox folder. Requires delegated Mail.ReadWrite.', [
          '复制邮件',
          '把邮件复制到文件夹',
          '保留邮件副本'
        ]),
        inputSchema: {
          messageId: z.string().min(1).max(300).describe('Message ID.'),
          destinationFolderId: z.string().min(1).max(300).describe('Destination folder ID.')
        }
      },
      async ({ messageId, destinationFolderId }) =>
        runTool(async () => copyMessage(config, messageId, destinationFolderId))
    );

    server.registerTool(
      'mail_delete_message',
      {
        title: 'Delete Mail',
        description: describeTool('Delete a message. Requires delegated Mail.ReadWrite.', [
          '删除邮件',
          '删掉这封邮件',
          '移除邮件'
        ]),
        inputSchema: {
          messageId: z.string().min(1).max(300).describe('Message ID.')
        }
      },
      async ({ messageId }) => runTool(async () => deleteMessage(config, messageId))
    );
  }
};
