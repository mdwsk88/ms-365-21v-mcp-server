import { z } from 'zod/v4';
import {
  createContactFolder,
  deleteContactFolder,
  getContact,
  listContactFolders,
  listFolderContacts,
  searchContacts,
  updateContactFolder
} from '../graph-extended.js';
import { createContact, deleteContact, listMyContacts, updateContact } from '../graph.js';
import { contactEmailSchema } from './schemas.js';
import { describeTool, runTool } from './results.js';
import type { ToolModule } from './types.js';

export const contactsModule: ToolModule = {
  category: 'contacts',
  displayName: 'Contacts',
  description: 'Mailbox contact read and write tools.',
  requiredRole: 'mcp.contacts',
  toolNames: [
    'contacts_list',
    'contacts_search',
    'contacts_get',
    'contacts_list_folders',
    'contacts_list_folder_contacts',
    'contacts_create_folder',
    'contacts_update_folder',
    'contacts_delete_folder',
    'contacts_create',
    'contacts_update',
    'contacts_delete'
  ],
  register(server, config) {
    server.registerTool(
      'contacts_list',
      {
        title: 'List My Contacts',
        description: describeTool('List mailbox contacts for the signed-in user. Requires delegated Contacts.Read.', [
          '查看我的联系人',
          '列出通讯录联系人',
          '查看邮箱联系人'
        ]),
        inputSchema: {
          top: z.number().int().min(1).max(100).optional().describe('Number of contacts to return, from 1 to 100.')
        }
      },
      async ({ top }) => runTool(async () => listMyContacts(config, top))
    );

    server.registerTool(
      'contacts_search',
      {
        title: 'Search My Contacts',
        description: describeTool('Search personal contacts by name prefix.', ['搜索联系人', '按姓名查联系人', '查找个人通讯录']),
        inputSchema: {
          query: z.string().min(1).max(200).describe('Name prefix to search.'),
          top: z.number().int().min(1).max(100).optional().describe('Number of contacts to return.')
        }
      },
      async ({ query, top }) => runTool(async () => searchContacts(config, query, top))
    );

    server.registerTool(
      'contacts_get',
      {
        title: 'Get Contact',
        description: describeTool('Read complete details for one personal contact.', ['查看联系人详情', '读取联系人信息']),
        inputSchema: {
          contactId: z.string().min(1).max(300).describe('Contact ID.')
        }
      },
      async ({ contactId }) => runTool(async () => getContact(config, contactId))
    );

    server.registerTool(
      'contacts_list_folders',
      {
        title: 'List Contact Folders',
        description: describeTool('List personal contact folders.', ['查看联系人文件夹', '列出通讯录分组']),
        inputSchema: {
          top: z.number().int().min(1).max(100).optional().describe('Number of folders to return.')
        }
      },
      async ({ top }) => runTool(async () => listContactFolders(config, top))
    );

    server.registerTool(
      'contacts_list_folder_contacts',
      {
        title: 'List Contacts In Folder',
        description: describeTool('List personal contacts in a selected contact folder.', [
          '查看联系人分组内容',
          '列出文件夹联系人',
          '查看通讯录分组'
        ]),
        inputSchema: {
          folderId: z.string().min(1).max(300).describe('Contact folder ID.'),
          top: z.number().int().min(1).max(100).optional().describe('Number of contacts to return.')
        }
      },
      async ({ folderId, top }) => runTool(async () => listFolderContacts(config, folderId, top))
    );

    server.registerTool(
      'contacts_create_folder',
      {
        title: 'Create Contact Folder',
        description: describeTool('Create a personal contact folder.', ['创建联系人文件夹', '新建通讯录分组']),
        inputSchema: {
          displayName: z.string().min(1).max(255).describe('Contact folder name.')
        }
      },
      async ({ displayName }) => runTool(async () => createContactFolder(config, displayName))
    );

    server.registerTool(
      'contacts_update_folder',
      {
        title: 'Rename Contact Folder',
        description: describeTool('Rename a personal contact folder.', ['重命名联系人文件夹', '修改通讯录分组名称']),
        inputSchema: {
          folderId: z.string().min(1).max(300).describe('Contact folder ID.'),
          displayName: z.string().min(1).max(255).describe('New folder name.')
        }
      },
      async ({ folderId, displayName }) => runTool(async () => updateContactFolder(config, folderId, displayName))
    );

    server.registerTool(
      'contacts_delete_folder',
      {
        title: 'Delete Contact Folder',
        description: describeTool('Delete a personal contact folder.', ['删除联系人文件夹', '移除通讯录分组']),
        inputSchema: {
          folderId: z.string().min(1).max(300).describe('Contact folder ID to delete.')
        }
      },
      async ({ folderId }) => runTool(async () => deleteContactFolder(config, folderId))
    );

    server.registerTool(
      'contacts_create',
      {
        title: 'Create Contact',
        description: describeTool(
          'Create a mailbox contact for the signed-in user. Requires delegated Contacts.ReadWrite.',
          ['创建联系人', '新增联系人', '添加邮箱联系人']
        ),
        inputSchema: {
          givenName: z.string().max(100).optional().describe('Given name.'),
          surname: z.string().max(100).optional().describe('Surname.'),
          displayName: z.string().max(200).optional().describe('Display name.'),
          emailAddresses: z.array(contactEmailSchema).max(10).optional().describe('Contact email addresses.'),
          businessPhones: z.array(z.string().min(1).max(80)).max(10).optional().describe('Business phone numbers.'),
          mobilePhone: z.string().max(80).optional().describe('Mobile phone number.'),
          companyName: z.string().max(200).optional().describe('Company name.'),
          jobTitle: z.string().max(200).optional().describe('Job title.')
        }
      },
      async (input) => runTool(async () => createContact(config, input))
    );

    server.registerTool(
      'contacts_update',
      {
        title: 'Update Contact',
        description: describeTool(
          'Update a mailbox contact for the signed-in user. Provide only the fields to change. Requires delegated Contacts.ReadWrite.',
          ['更新联系人', '修改联系人', '编辑联系人信息']
        ),
        inputSchema: {
          contactId: z.string().min(1).max(300).describe('Contact ID.'),
          givenName: z.string().max(100).optional().describe('Given name.'),
          surname: z.string().max(100).optional().describe('Surname.'),
          displayName: z.string().max(200).optional().describe('Display name.'),
          emailAddresses: z.array(contactEmailSchema).max(10).optional().describe('Contact email addresses.'),
          businessPhones: z.array(z.string().min(1).max(80)).max(10).optional().describe('Business phone numbers.'),
          mobilePhone: z.string().max(80).optional().describe('Mobile phone number.'),
          companyName: z.string().max(200).optional().describe('Company name.'),
          jobTitle: z.string().max(200).optional().describe('Job title.')
        }
      },
      async ({ contactId, ...input }) => runTool(async () => updateContact(config, contactId, input))
    );

    server.registerTool(
      'contacts_delete',
      {
        title: 'Delete Contact',
        description: describeTool(
          'Delete a mailbox contact for the signed-in user. Requires delegated Contacts.ReadWrite.',
          ['删除联系人', '移除联系人', '删掉邮箱联系人']
        ),
        inputSchema: {
          contactId: z.string().min(1).max(300).describe('Contact ID.')
        }
      },
      async ({ contactId }) => runTool(async () => deleteContact(config, contactId))
    );
  }
};
