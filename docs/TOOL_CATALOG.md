# MS 365-21V MCP Server - Tool Catalog

> This file is generated from the runtime registry by `npm run docs:tools`. Do not edit tool rows manually.

The implementation contains **150 tools**. The current production permission boundary exposes **125 tools** and hides **25 tools** whose Graph delegated scopes are not approved. Every enabled business tool is mapped to an Entra App Role and one or more Microsoft Graph delegated permissions.

## Summary

| Module | Tools | Read | Write | App Role |
|---|---:|---:|---:|---|
| Calendar | 17 | 7 | 10 | `mcp.calendar` |
| Contacts | 11 | 5 | 6 | `mcp.contacts` |
| OneDrive | 18 | 9 | 9 | `mcp.drive` |
| Gateway Authentication | 2 | 2 | 0 | - |
| Groups | 1 | 1 | 0 | `mcp.groups` |
| Mail | 23 | 9 | 14 | `mcp.mail` |
| Microsoft Search | 5 | 5 | 0 | `mcp.calendar`, `mcp.drive`, `mcp.mail`, `mcp.search`, `mcp.sharepoint`, `mcp.teams` |
| SharePoint | 33 | 15 | 18 | `mcp.sharepoint` |
| Smart M365 Insights | 2 | 2 | 0 | `mcp.calendar`, `mcp.mail`, `mcp.smart` |
| Teams | 10 | 7 | 3 | `mcp.teams` |
| Users and Profiles | 3 | 3 | 0 | `mcp.users` |

## Calendar

| Tool | Title | Access | Required App Roles | Graph delegated scopes |
|---|---|---|---|---|
| `calendar_add_file_attachment` | Add File Attachment To Event | Write | `mcp.calendar` | `Calendars.ReadWrite` |
| `calendar_cancel_event` | Cancel Calendar Event | Write | `mcp.calendar` | `Calendars.ReadWrite` |
| `calendar_create_calendar` | Create Calendar | Write | `mcp.calendar` | `Calendars.ReadWrite` |
| `calendar_create_event` | Create Calendar Event | Write | `mcp.calendar` | `Calendars.ReadWrite` |
| `calendar_delete_attachment` | Delete Calendar Event Attachment | Write | `mcp.calendar` | `Calendars.ReadWrite` |
| `calendar_delete_calendar` | Delete Calendar | Write | `mcp.calendar` | `Calendars.ReadWrite` |
| `calendar_delete_event` | Delete Calendar Event | Write | `mcp.calendar` | `Calendars.ReadWrite` |
| `calendar_get_event` | Get Calendar Event | Read | `mcp.calendar` | `Calendars.Read` |
| `calendar_get_schedule` | Get Free Busy Schedule | Read | `mcp.calendar` | `Calendars.Read` |
| `calendar_list_attachments` | List Calendar Event Attachments | Read | `mcp.calendar` | `Calendars.Read` |
| `calendar_list_calendars` | List My Calendars | Read | `mcp.calendar` | `Calendars.Read` |
| `calendar_list_event_instances` | List Recurring Event Instances | Read | `mcp.calendar` | `Calendars.Read` |
| `calendar_list_events` | List My Calendar Events | Read | `mcp.calendar` | `Calendars.Read` |
| `calendar_list_view` | List Calendar View | Read | `mcp.calendar` | `Calendars.Read` |
| `calendar_respond_to_event` | Respond To Calendar Event | Write | `mcp.calendar` | `Calendars.ReadWrite` |
| `calendar_update_calendar` | Rename Calendar | Write | `mcp.calendar` | `Calendars.ReadWrite` |
| `calendar_update_event` | Update Calendar Event | Write | `mcp.calendar` | `Calendars.ReadWrite` |

## Contacts

| Tool | Title | Access | Required App Roles | Graph delegated scopes |
|---|---|---|---|---|
| `contacts_create` | Create Contact | Write | `mcp.contacts` | `Contacts.ReadWrite` |
| `contacts_create_folder` | Create Contact Folder | Write | `mcp.contacts` | `Contacts.ReadWrite` |
| `contacts_delete` | Delete Contact | Write | `mcp.contacts` | `Contacts.ReadWrite` |
| `contacts_delete_folder` | Delete Contact Folder | Write | `mcp.contacts` | `Contacts.ReadWrite` |
| `contacts_get` | Get Contact | Read | `mcp.contacts` | `Contacts.Read` |
| `contacts_list` | List My Contacts | Read | `mcp.contacts` | `Contacts.Read` |
| `contacts_list_folder_contacts` | List Contacts In Folder | Read | `mcp.contacts` | `Contacts.Read` |
| `contacts_list_folders` | List Contact Folders | Read | `mcp.contacts` | `Contacts.Read` |
| `contacts_search` | Search My Contacts | Read | `mcp.contacts` | `Contacts.Read` |
| `contacts_update` | Update Contact | Write | `mcp.contacts` | `Contacts.ReadWrite` |
| `contacts_update_folder` | Rename Contact Folder | Write | `mcp.contacts` | `Contacts.ReadWrite` |

## OneDrive

| Tool | Title | Access | Required App Roles | Graph delegated scopes |
|---|---|---|---|---|
| `drive_copy_item` | Copy OneDrive Item | Write | `mcp.drive` | `Files.ReadWrite` |
| `drive_create_folder` | Create OneDrive Folder | Write | `mcp.drive` | `Files.ReadWrite` |
| `drive_create_share_link` | Create OneDrive Share Link | Write | `mcp.drive` | `Files.ReadWrite` |
| `drive_delete_item` | Delete OneDrive Item | Write | `mcp.drive` | `Files.ReadWrite` |
| `drive_download_file` | Download OneDrive File | Read | `mcp.drive` | `Files.Read` |
| `drive_get_drive` | Get My OneDrive | Read | `mcp.drive` | `Files.Read` |
| `drive_get_item` | Get OneDrive Item | Read | `mcp.drive` | `Files.Read` |
| `drive_invite_item` | Invite People To OneDrive Item | Write | `mcp.drive` | `Files.ReadWrite` |
| `drive_list_children` | List OneDrive Folder Items | Read | `mcp.drive` | `Files.Read` |
| `drive_list_permissions` | List OneDrive Item Permissions | Read | `mcp.drive` | `Files.Read` |
| `drive_list_recent` | List Recent OneDrive Items | Read | `mcp.drive` | `Files.Read` |
| `drive_list_root` | List My OneDrive Root | Read | `mcp.drive` | `Files.Read` |
| `drive_list_versions` | List OneDrive File Versions | Read | `mcp.drive` | `Files.Read` |
| `drive_move_item` | Move OneDrive Item | Write | `mcp.drive` | `Files.ReadWrite` |
| `drive_rename_item` | Rename OneDrive Item | Write | `mcp.drive` | `Files.ReadWrite` |
| `drive_restore_version` | Restore OneDrive File Version | Write | `mcp.drive` | `Files.ReadWrite` |
| `drive_search_items` | Search My OneDrive | Read | `mcp.drive` | `Files.Read` |
| `drive_upload_small_file` | Upload Small OneDrive File | Write | `mcp.drive` | `Files.ReadWrite` |

## Gateway Authentication

| Tool | Title | Access | Required App Roles | Graph delegated scopes |
|---|---|---|---|---|
| `auth_status` | Check Auth Status | Read | - | - |
| `confirm_execute` | Execute Approved Operation | Read | - | - |

## Groups

| Tool | Title | Access | Required App Roles | Graph delegated scopes |
|---|---|---|---|---|
| `groups_check_my_memberships` | Check My Group Memberships | Read | `mcp.groups` | `User.Read` |

## Mail

| Tool | Title | Access | Required App Roles | Graph delegated scopes |
|---|---|---|---|---|
| `mail_add_file_attachment` | Add File Attachment To Mail | Write | `mcp.mail` | `Mail.ReadWrite` |
| `mail_copy_message` | Copy Mail | Write | `mcp.mail` | `Mail.ReadWrite` |
| `mail_create_draft` | Create Mail Draft | Write | `mcp.mail` | `Mail.ReadWrite` |
| `mail_create_folder` | Create Mail Folder | Write | `mcp.mail` | `Mail.ReadWrite` |
| `mail_delete_attachment` | Delete Mail Attachment | Write | `mcp.mail` | `Mail.ReadWrite` |
| `mail_delete_folder` | Delete Mail Folder | Write | `mcp.mail` | `Mail.ReadWrite` |
| `mail_delete_message` | Delete Mail | Write | `mcp.mail` | `Mail.ReadWrite` |
| `mail_forward` | Forward Mail | Write | `mcp.mail` | `Mail.Send` |
| `mail_get_attachment` | Get Mail Attachment | Read | `mcp.mail` | `Mail.Read` |
| `mail_get_folder` | Get Mail Folder | Read | `mcp.mail` | `Mail.Read` |
| `mail_get_message` | Get Mail Details | Read | `mcp.mail` | `Mail.Read` |
| `mail_list_attachments` | List Mail Attachments | Read | `mcp.mail` | `Mail.Read` |
| `mail_list_child_folders` | List Child Mail Folders | Read | `mcp.mail` | `Mail.Read` |
| `mail_list_folder_messages` | List Mail In Folder | Read | `mcp.mail` | `Mail.Read` |
| `mail_list_folders` | List Mail Folders | Read | `mcp.mail` | `Mail.Read` |
| `mail_list_messages` | List My Mail | Read | `mcp.mail` | `Mail.Read` |
| `mail_move_message` | Move Mail | Write | `mcp.mail` | `Mail.ReadWrite` |
| `mail_reply` | Reply To Mail | Write | `mcp.mail` | `Mail.Send` |
| `mail_search_messages` | Search My Mail | Read | `mcp.mail` | `Mail.Read` |
| `mail_send` | Send Mail | Write | `mcp.mail` | `Mail.Send` |
| `mail_send_draft` | Send Mail Draft | Write | `mcp.mail` | `Mail.Send` |
| `mail_set_read_state` | Set Mail Read State | Write | `mcp.mail` | `Mail.ReadWrite` |
| `mail_update_folder` | Rename Mail Folder | Write | `mcp.mail` | `Mail.ReadWrite` |

## Microsoft Search

| Tool | Title | Access | Required App Roles | Graph delegated scopes |
|---|---|---|---|---|
| `search_calendar` | Search Microsoft 365 Calendar | Read | `mcp.search`, `mcp.calendar` | `Calendars.Read` |
| `search_files` | Search Microsoft 365 Files | Read | `mcp.search`, `mcp.drive` | `Files.Read.All` |
| `search_mail` | Search Microsoft 365 Mail | Read | `mcp.search`, `mcp.mail` | `Mail.Read` |
| `search_sharepoint` | Search SharePoint Content | Read | `mcp.search`, `mcp.sharepoint` | `Sites.Read.All` |
| `search_teams` | Search Teams Messages | Read | `mcp.search`, `mcp.teams` | `Chat.Read` |

## SharePoint

| Tool | Title | Access | Required App Roles | Graph delegated scopes |
|---|---|---|---|---|
| `sharepoint_copy_drive_item` | Copy SharePoint Library Item | Write | `mcp.sharepoint` | `Files.ReadWrite.All` |
| `sharepoint_create_column` | Create SharePoint List Column | Write | `mcp.sharepoint` | `Sites.Manage.All` |
| `sharepoint_create_drive_folder` | Create SharePoint Library Folder | Write | `mcp.sharepoint` | `Sites.ReadWrite.All`, `Files.ReadWrite.All` |
| `sharepoint_create_list` | Create SharePoint List | Write | `mcp.sharepoint` | `Sites.Manage.All` |
| `sharepoint_create_list_item` | Create SharePoint List Item | Write | `mcp.sharepoint` | `Sites.ReadWrite.All` |
| `sharepoint_create_share_link` | Create SharePoint Share Link | Write | `mcp.sharepoint` | `Files.ReadWrite.All` |
| `sharepoint_delete_column` | Delete SharePoint List Column | Write | `mcp.sharepoint` | `Sites.Manage.All` |
| `sharepoint_delete_drive_item` | Delete SharePoint Library Item | Write | `mcp.sharepoint` | `Sites.ReadWrite.All`, `Files.ReadWrite.All` |
| `sharepoint_delete_list` | Delete SharePoint List | Write | `mcp.sharepoint` | `Sites.Manage.All` |
| `sharepoint_delete_list_item` | Delete SharePoint List Item | Write | `mcp.sharepoint` | `Sites.ReadWrite.All` |
| `sharepoint_download_file` | Download SharePoint File | Read | `mcp.sharepoint` | `Files.Read.All` |
| `sharepoint_get_drive_item` | Get SharePoint Library Item | Read | `mcp.sharepoint` | `Sites.Read.All`, `Files.Read.All` |
| `sharepoint_get_list_item` | Get SharePoint List Item | Read | `mcp.sharepoint` | `Sites.Read.All` |
| `sharepoint_get_site` | Get SharePoint Site | Read | `mcp.sharepoint` | `Sites.Read.All` |
| `sharepoint_get_site_by_path` | Get SharePoint Site By Path | Read | `mcp.sharepoint` | `Sites.Read.All` |
| `sharepoint_invite_drive_item` | Invite People To SharePoint Item | Write | `mcp.sharepoint` | `Files.ReadWrite.All` |
| `sharepoint_list_columns` | List SharePoint List Columns | Read | `mcp.sharepoint` | `Sites.Read.All` |
| `sharepoint_list_drive_items` | List SharePoint Library Items | Read | `mcp.sharepoint` | `Sites.Read.All`, `Files.Read.All` |
| `sharepoint_list_drive_permissions` | List SharePoint Item Permissions | Read | `mcp.sharepoint` | `Files.Read.All` |
| `sharepoint_list_drive_versions` | List SharePoint File Versions | Read | `mcp.sharepoint` | `Files.Read.All` |
| `sharepoint_list_drives` | List SharePoint Document Libraries | Read | `mcp.sharepoint` | `Sites.Read.All`, `Files.Read.All` |
| `sharepoint_list_item_delta` | Read SharePoint List Changes | Read | `mcp.sharepoint` | `Sites.Read.All` |
| `sharepoint_list_list_items` | List SharePoint List Items | Read | `mcp.sharepoint` | `Sites.Read.All` |
| `sharepoint_list_lists` | List SharePoint Lists | Read | `mcp.sharepoint` | `Sites.Read.All` |
| `sharepoint_move_drive_item` | Move SharePoint Library Item | Write | `mcp.sharepoint` | `Files.ReadWrite.All` |
| `sharepoint_rename_drive_item` | Rename SharePoint Library Item | Write | `mcp.sharepoint` | `Files.ReadWrite.All` |
| `sharepoint_restore_drive_version` | Restore SharePoint File Version | Write | `mcp.sharepoint` | `Files.ReadWrite.All` |
| `sharepoint_search_drive_items` | Search SharePoint Library | Read | `mcp.sharepoint` | `Sites.Read.All`, `Files.Read.All` |
| `sharepoint_search_sites` | Search SharePoint Sites | Read | `mcp.sharepoint` | `Sites.Read.All` |
| `sharepoint_update_column` | Update SharePoint List Column | Write | `mcp.sharepoint` | `Sites.Manage.All` |
| `sharepoint_update_list` | Update SharePoint List | Write | `mcp.sharepoint` | `Sites.Manage.All` |
| `sharepoint_update_list_item` | Update SharePoint List Item | Write | `mcp.sharepoint` | `Sites.ReadWrite.All` |
| `sharepoint_upload_small_file` | Upload Small SharePoint File | Write | `mcp.sharepoint` | `Sites.ReadWrite.All`, `Files.ReadWrite.All` |

## Smart M365 Insights

| Tool | Title | Access | Required App Roles | Graph delegated scopes |
|---|---|---|---|---|
| `smart_calendar_conflicts` | Find Calendar Conflicts | Read | `mcp.smart`, `mcp.calendar` | `Calendars.Read` |
| `smart_mail_digest` | Build Mail Digest | Read | `mcp.smart`, `mcp.mail` | `Mail.Read`, `User.Read` |

## Teams

| Tool | Title | Access | Required App Roles | Graph delegated scopes |
|---|---|---|---|---|
| `teams_create_chat` | Create Teams Chat | Write | `mcp.teams` | `Chat.Create` |
| `teams_delete_chat_message` | Delete Teams Chat Message | Write | `mcp.teams` | `Chat.ReadWrite` |
| `teams_get_channel` | Get Teams Channel | Read | `mcp.teams` | `Channel.ReadBasic.All` |
| `teams_get_team` | Get Teams Team | Read | `mcp.teams` | `Team.ReadBasic.All` |
| `teams_list_channels` | List Team Channels | Read | `mcp.teams` | `Channel.ReadBasic.All` |
| `teams_list_chat_members` | List Teams Chat Members | Read | `mcp.teams` | `Chat.Read` |
| `teams_list_chat_messages` | List Teams Chat Messages | Read | `mcp.teams` | `Chat.Read` |
| `teams_list_chats` | List My Teams Chats | Read | `mcp.teams` | `Chat.Read` |
| `teams_list_joined_teams` | List My Joined Teams | Read | `mcp.teams` | `Team.ReadBasic.All` |
| `teams_send_chat_message` | Send Teams Chat Message | Write | `mcp.teams` | `ChatMessage.Send` |

## Users and Profiles

| Tool | Title | Access | Required App Roles | Graph delegated scopes |
|---|---|---|---|---|
| `graph_get_me` | Get My Profile | Read | `mcp.users` | `User.Read` |
| `users_list` | List Organization Users | Read | `mcp.users` | `User.ReadBasic.All` |
| `users_search` | Search Organization Users | Read | `mcp.users` | `User.ReadBasic.All` |

## Disabled Pending Tenant Approval

The following **25 tools** remain implemented but are excluded from both MCP discovery and direct invocation. They become available after the corresponding delegated permissions are approved and removed from `MCP_DISABLED_GRAPH_SCOPES`.

| Tool | Module | Graph delegated scopes |
|---|---|---|
| `groups_get` | Groups | `GroupMember.Read.All` |
| `groups_list` | Groups | `GroupMember.Read.All` |
| `groups_list_members` | Groups | `GroupMember.Read.All` |
| `groups_list_owners` | Groups | `GroupMember.Read.All` |
| `smart_teams_unread` | Smart M365 Insights | `Team.ReadBasic.All`, `Channel.ReadBasic.All`, `ChannelMessage.Read.All`, `Chat.Read` |
| `teams_add_channel_member` | Teams | `ChannelMember.ReadWrite.All` |
| `teams_add_chat_member` | Teams | `ChatMember.ReadWrite` |
| `teams_add_team_member` | Teams | `TeamMember.ReadWrite.All` |
| `teams_create_channel` | Teams | `Channel.Create` |
| `teams_delete_channel` | Teams | `Channel.Delete.All` |
| `teams_delete_channel_message` | Teams | `ChannelMessage.ReadWrite` |
| `teams_list_channel_members` | Teams | `ChannelMember.Read.All` |
| `teams_list_channel_message_replies` | Teams | `ChannelMessage.Read.All` |
| `teams_list_channel_messages` | Teams | `ChannelMessage.Read.All` |
| `teams_list_team_members` | Teams | `TeamMember.Read.All` |
| `teams_remove_channel_member` | Teams | `ChannelMember.ReadWrite.All` |
| `teams_remove_chat_member` | Teams | `ChatMember.ReadWrite` |
| `teams_remove_team_member` | Teams | `TeamMember.ReadWrite.All` |
| `teams_reply_channel_message` | Teams | `ChannelMessage.Send` |
| `teams_send_channel_message` | Teams | `ChannelMessage.Send` |
| `teams_update_channel` | Teams | `ChannelSettings.ReadWrite.All` |
| `users_get_manager` | Users and Profiles | `User.Read.All` |
| `users_get_profile` | Users and Profiles | `User.Read.All` |
| `users_list_direct_reports` | Users and Profiles | `User.Read.All` |
| `users_list_memberships` | Users and Profiles | `User.Read.All` |

## Coverage Boundary

The catalog targets Microsoft Graph v1.0 APIs that are available in the China cloud operated by 21Vianet, support delegated work-or-school identities, and are appropriate for interactive MCP use. It intentionally excludes beta-only APIs, application-only tenant administration, subscriptions/webhooks, security and Intune administration, and workloads whose Microsoft documentation marks China as unsupported.
