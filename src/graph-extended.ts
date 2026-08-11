import type { AppConfig } from './config.js';
import { graphRequest } from './graph.js';

type RecipientInput = { email: string; name?: string };
type EventDateTimeInput = { dateTime: string; timeZone?: string };
type ConversationMemberInput = { userId: string; role?: 'owner' | 'guest' | 'member' };

function segment(value: string): string {
  return encodeURIComponent(value);
}

function clampTop(top: number | undefined, fallback: number, max: number): number {
  if (!Number.isFinite(top)) return fallback;
  return Math.max(1, Math.min(max, Math.trunc(top as number)));
}

function escapeODataString(value: string): string {
  return value.replace(/'/g, "''");
}

function encodeDrivePath(value: string): string {
  return value
    .replace(/^\/+|\/+$/g, '')
    .split('/')
    .filter(Boolean)
    .map(segment)
    .join('/');
}

function encodeSitePath(value: string): string {
  const encoded = encodeDrivePath(value);
  return encoded ? `/${encoded}` : '/';
}

function eventDateTime(value: EventDateTimeInput) {
  return { dateTime: value.dateTime, timeZone: value.timeZone ?? 'China Standard Time' };
}

function recipient(value: RecipientInput) {
  return { emailAddress: { address: value.email, name: value.name } };
}

function userBinding(config: AppConfig, userId: string): string {
  return `${config.graphBaseUrl}/users('${escapeODataString(userId)}')`;
}

function conversationMember(
  config: AppConfig,
  member: ConversationMemberInput,
  defaultRole: 'member' | 'owner' = 'member'
) {
  const role = member.role ?? defaultRole;
  const roles = role === 'member' ? [] : [role];
  return {
    '@odata.type': '#microsoft.graph.aadUserConversationMember',
    'user@odata.bind': userBinding(config, member.userId),
    roles
  };
}

// Organization users and groups

export async function listOrganizationUsers(config: AppConfig, top?: number): Promise<unknown> {
  return graphRequest(config, 'GET', '/users', {
    query: {
      $top: clampTop(top, 25, 100),
      $select: 'id,displayName,userPrincipalName,mail,jobTitle,department,officeLocation,accountEnabled',
      $orderby: 'displayName'
    },
    scopes: ['User.ReadBasic.All']
  });
}

export async function getUserManager(config: AppConfig, userId = 'me'): Promise<unknown> {
  const path = userId === 'me' ? '/me/manager' : `/users/${segment(userId)}/manager`;
  return graphRequest(config, 'GET', path, {
    query: { $select: 'id,displayName,userPrincipalName,mail,jobTitle,department,officeLocation' },
    scopes: ['User.Read.All']
  });
}

export async function listUserDirectReports(config: AppConfig, userId = 'me', top?: number): Promise<unknown> {
  const path = userId === 'me' ? '/me/directReports' : `/users/${segment(userId)}/directReports`;
  return graphRequest(config, 'GET', path, {
    query: {
      $top: clampTop(top, 25, 100),
      $select: 'id,displayName,userPrincipalName,mail,jobTitle,department,officeLocation'
    },
    scopes: ['User.Read.All']
  });
}

export async function listUserMemberships(config: AppConfig, userId = 'me', top?: number): Promise<unknown> {
  const path = userId === 'me' ? '/me/memberOf' : `/users/${segment(userId)}/memberOf`;
  return graphRequest(config, 'GET', path, {
    query: {
      $top: clampTop(top, 50, 100),
      $select: 'id,displayName,description,mail,mailEnabled,securityEnabled,groupTypes'
    },
    scopes: userId === 'me' ? ['User.Read'] : ['User.Read.All']
  });
}

export async function listGroups(config: AppConfig, query?: string, top?: number): Promise<unknown> {
  const filter = query?.trim()
    ? `startswith(displayName,'${escapeODataString(query.trim())}')`
    : undefined;
  return graphRequest(config, 'GET', '/groups', {
    query: {
      $top: clampTop(top, 25, 100),
      $select: 'id,displayName,description,mail,mailEnabled,securityEnabled,visibility,groupTypes',
      $filter: filter,
      $orderby: filter ? undefined : 'displayName'
    },
    scopes: ['GroupMember.Read.All']
  });
}

export async function getGroup(config: AppConfig, groupId: string): Promise<unknown> {
  return graphRequest(config, 'GET', `/groups/${segment(groupId)}`, {
    query: {
      $select: 'id,displayName,description,mail,mailEnabled,securityEnabled,visibility,groupTypes,createdDateTime'
    },
    scopes: ['GroupMember.Read.All']
  });
}

export async function listGroupMembers(config: AppConfig, groupId: string, top?: number): Promise<unknown> {
  return graphRequest(config, 'GET', `/groups/${segment(groupId)}/members`, {
    query: {
      $top: clampTop(top, 50, 100),
      $select: 'id,displayName,userPrincipalName,mail,jobTitle,department'
    },
    scopes: ['GroupMember.Read.All']
  });
}

export async function listGroupOwners(config: AppConfig, groupId: string, top?: number): Promise<unknown> {
  return graphRequest(config, 'GET', `/groups/${segment(groupId)}/owners`, {
    query: {
      $top: clampTop(top, 25, 100),
      $select: 'id,displayName,userPrincipalName,mail,jobTitle,department'
    },
    scopes: ['GroupMember.Read.All']
  });
}

export async function checkMyGroupMemberships(config: AppConfig, groupIds: string[]): Promise<unknown> {
  return graphRequest(config, 'POST', '/me/checkMemberGroups', {
    body: { groupIds },
    scopes: ['User.Read']
  });
}

// Outlook mail additions

export async function listFolderMessages(
  config: AppConfig,
  folderId: string,
  top?: number,
  unreadOnly?: boolean
): Promise<unknown> {
  return graphRequest(config, 'GET', `/me/mailFolders/${segment(folderId)}/messages`, {
    query: {
      $top: clampTop(top, 20, 50),
      $select: 'id,subject,from,toRecipients,receivedDateTime,webLink,isRead,importance,hasAttachments',
      $orderby: 'receivedDateTime desc',
      $filter: unreadOnly ? 'isRead eq false' : undefined
    },
    scopes: ['Mail.Read']
  });
}

export async function getMailFolder(config: AppConfig, folderId: string): Promise<unknown> {
  return graphRequest(config, 'GET', `/me/mailFolders/${segment(folderId)}`, {
    query: { $select: 'id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount,isHidden' },
    scopes: ['Mail.Read']
  });
}

export async function listChildMailFolders(
  config: AppConfig,
  folderId: string,
  top?: number,
  includeHidden?: boolean
): Promise<unknown> {
  return graphRequest(config, 'GET', `/me/mailFolders/${segment(folderId)}/childFolders`, {
    query: {
      $top: clampTop(top, 25, 100),
      includeHiddenFolders: includeHidden,
      $select: 'id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount,isHidden'
    },
    scopes: ['Mail.Read']
  });
}

export async function createMailFolder(
  config: AppConfig,
  displayName: string,
  parentFolderId?: string,
  isHidden = false
): Promise<unknown> {
  const path = parentFolderId
    ? `/me/mailFolders/${segment(parentFolderId)}/childFolders`
    : '/me/mailFolders';
  return graphRequest(config, 'POST', path, {
    body: { displayName, isHidden },
    scopes: ['Mail.ReadWrite']
  });
}

export async function updateMailFolder(config: AppConfig, folderId: string, displayName: string): Promise<unknown> {
  return graphRequest(config, 'PATCH', `/me/mailFolders/${segment(folderId)}`, {
    body: { displayName },
    scopes: ['Mail.ReadWrite']
  });
}

export async function deleteMailFolder(config: AppConfig, folderId: string): Promise<unknown> {
  return graphRequest(config, 'DELETE', `/me/mailFolders/${segment(folderId)}`, {
    scopes: ['Mail.ReadWrite']
  });
}

export async function copyMessage(config: AppConfig, messageId: string, destinationFolderId: string): Promise<unknown> {
  return graphRequest(config, 'POST', `/me/messages/${segment(messageId)}/copy`, {
    body: { destinationId: destinationFolderId },
    scopes: ['Mail.ReadWrite']
  });
}

export async function getMessageAttachment(
  config: AppConfig,
  messageId: string,
  attachmentId: string,
  raw = false
): Promise<unknown> {
  const path = `/me/messages/${segment(messageId)}/attachments/${segment(attachmentId)}${raw ? '/$value' : ''}`;
  return graphRequest(config, 'GET', path, {
    query: raw ? undefined : { $expand: 'microsoft.graph.itemattachment/item' },
    scopes: ['Mail.Read'],
    responseType: raw ? 'base64' : 'auto',
    maxResponseBytes: 10 * 1024 * 1024
  });
}

export async function addMessageFileAttachment(
  config: AppConfig,
  messageId: string,
  name: string,
  contentType: string,
  contentBase64: string,
  isInline = false,
  contentId?: string
): Promise<unknown> {
  return graphRequest(config, 'POST', `/me/messages/${segment(messageId)}/attachments`, {
    body: {
      '@odata.type': '#microsoft.graph.fileAttachment',
      name,
      contentType,
      contentBytes: contentBase64,
      isInline,
      contentId
    },
    scopes: ['Mail.ReadWrite']
  });
}

export async function deleteMessageAttachment(
  config: AppConfig,
  messageId: string,
  attachmentId: string
): Promise<unknown> {
  return graphRequest(config, 'DELETE', `/me/messages/${segment(messageId)}/attachments/${segment(attachmentId)}`, {
    scopes: ['Mail.ReadWrite']
  });
}

// Calendar additions

export async function listCalendarView(
  config: AppConfig,
  start: string,
  end: string,
  top?: number,
  timeZone = 'China Standard Time'
): Promise<unknown> {
  return graphRequest(config, 'GET', '/me/calendarView', {
    query: {
      startDateTime: start,
      endDateTime: end,
      $top: clampTop(top, 25, 100),
      $select: 'id,subject,organizer,attendees,start,end,location,webLink,isOnlineMeeting,responseStatus,seriesMasterId',
      $orderby: 'start/dateTime'
    },
    prefer: `outlook.timezone="${timeZone.replace(/"/g, '')}"`,
    scopes: ['Calendars.Read']
  });
}

export async function listCalendarEventInstances(
  config: AppConfig,
  eventId: string,
  start: string,
  end: string,
  top?: number
): Promise<unknown> {
  return graphRequest(config, 'GET', `/me/events/${segment(eventId)}/instances`, {
    query: {
      startDateTime: start,
      endDateTime: end,
      $top: clampTop(top, 25, 100),
      $select: 'id,subject,organizer,attendees,start,end,location,webLink,responseStatus,seriesMasterId'
    },
    prefer: 'outlook.timezone="China Standard Time"',
    scopes: ['Calendars.Read']
  });
}

export async function createCalendar(config: AppConfig, name: string): Promise<unknown> {
  return graphRequest(config, 'POST', '/me/calendars', {
    body: { name },
    scopes: ['Calendars.ReadWrite']
  });
}

export async function updateCalendar(config: AppConfig, calendarId: string, name: string): Promise<unknown> {
  return graphRequest(config, 'PATCH', `/me/calendars/${segment(calendarId)}`, {
    body: { name },
    scopes: ['Calendars.ReadWrite']
  });
}

export async function deleteCalendar(config: AppConfig, calendarId: string): Promise<unknown> {
  return graphRequest(config, 'DELETE', `/me/calendars/${segment(calendarId)}`, {
    scopes: ['Calendars.ReadWrite']
  });
}

export async function getSchedule(
  config: AppConfig,
  schedules: string[],
  start: EventDateTimeInput,
  end: EventDateTimeInput,
  availabilityViewInterval?: number
): Promise<unknown> {
  return graphRequest(config, 'POST', '/me/calendar/getSchedule', {
    body: {
      schedules,
      startTime: eventDateTime(start),
      endTime: eventDateTime(end),
      availabilityViewInterval
    },
    prefer: `outlook.timezone="${(start.timeZone ?? 'China Standard Time').replace(/"/g, '')}"`,
    scopes: ['Calendars.Read']
  });
}

export async function cancelCalendarEvent(config: AppConfig, eventId: string, comment?: string): Promise<unknown> {
  return graphRequest(config, 'POST', `/me/events/${segment(eventId)}/cancel`, {
    body: { comment: comment ?? '' },
    scopes: ['Calendars.ReadWrite']
  });
}

export async function listEventAttachments(config: AppConfig, eventId: string): Promise<unknown> {
  return graphRequest(config, 'GET', `/me/events/${segment(eventId)}/attachments`, {
    query: { $select: 'id,name,contentType,size,isInline,lastModifiedDateTime' },
    scopes: ['Calendars.Read']
  });
}

export async function addEventFileAttachment(
  config: AppConfig,
  eventId: string,
  name: string,
  contentType: string,
  contentBase64: string,
  isInline = false,
  contentId?: string
): Promise<unknown> {
  return graphRequest(config, 'POST', `/me/events/${segment(eventId)}/attachments`, {
    body: {
      '@odata.type': '#microsoft.graph.fileAttachment',
      name,
      contentType,
      contentBytes: contentBase64,
      isInline,
      contentId
    },
    scopes: ['Calendars.ReadWrite']
  });
}

export async function deleteEventAttachment(
  config: AppConfig,
  eventId: string,
  attachmentId: string
): Promise<unknown> {
  return graphRequest(config, 'DELETE', `/me/events/${segment(eventId)}/attachments/${segment(attachmentId)}`, {
    scopes: ['Calendars.ReadWrite']
  });
}

// OneDrive additions

export async function getMyDrive(config: AppConfig): Promise<unknown> {
  return graphRequest(config, 'GET', '/me/drive', {
    query: { $select: 'id,name,description,driveType,webUrl,createdDateTime,lastModifiedDateTime,owner,quota' },
    scopes: ['Files.Read']
  });
}

export async function listDriveChildren(config: AppConfig, parentItemId: string, top?: number): Promise<unknown> {
  return graphRequest(config, 'GET', `/me/drive/items/${segment(parentItemId)}/children`, {
    query: {
      $top: clampTop(top, 25, 100),
      $select: 'id,name,webUrl,size,createdDateTime,lastModifiedDateTime,folder,file,parentReference'
    },
    scopes: ['Files.Read']
  });
}

export async function downloadDriveFile(config: AppConfig, itemId: string): Promise<unknown> {
  return graphRequest(config, 'GET', `/me/drive/items/${segment(itemId)}/content`, {
    scopes: ['Files.Read'],
    responseType: 'base64',
    maxResponseBytes: 10 * 1024 * 1024
  });
}

export async function copyDriveItem(
  config: AppConfig,
  itemId: string,
  parentItemId: string,
  newName?: string
): Promise<unknown> {
  return graphRequest(config, 'POST', `/me/drive/items/${segment(itemId)}/copy`, {
    body: { parentReference: { id: parentItemId }, name: newName },
    scopes: ['Files.ReadWrite']
  });
}

export async function listDriveVersions(config: AppConfig, itemId: string): Promise<unknown> {
  return graphRequest(config, 'GET', `/me/drive/items/${segment(itemId)}/versions`, {
    scopes: ['Files.Read']
  });
}

export async function restoreDriveVersion(config: AppConfig, itemId: string, versionId: string): Promise<unknown> {
  return graphRequest(config, 'POST', `/me/drive/items/${segment(itemId)}/versions/${segment(versionId)}/restoreVersion`, {
    scopes: ['Files.ReadWrite']
  });
}

export async function inviteDriveItem(
  config: AppConfig,
  itemId: string,
  recipientsInput: RecipientInput[],
  roles: Array<'read' | 'write'>,
  message?: string,
  requireSignIn = true,
  sendInvitation = true
): Promise<unknown> {
  return graphRequest(config, 'POST', `/me/drive/items/${segment(itemId)}/invite`, {
    body: {
      recipients: recipientsInput.map(recipient),
      roles,
      message,
      requireSignIn,
      sendInvitation
    },
    scopes: ['Files.ReadWrite']
  });
}

export async function listRecentDriveItems(config: AppConfig, top?: number): Promise<unknown> {
  return graphRequest(config, 'GET', '/me/drive/recent', {
    query: {
      $top: clampTop(top, 25, 100),
      $select: 'id,name,webUrl,size,lastModifiedDateTime,folder,file,parentReference,remoteItem'
    },
    scopes: ['Files.Read']
  });
}

// SharePoint additions

export async function getSharePointSiteByPath(
  config: AppConfig,
  hostname: string,
  relativePath: string
): Promise<unknown> {
  return graphRequest(config, 'GET', `/sites/${hostname}:${encodeSitePath(relativePath)}`, {
    query: { $select: 'id,name,displayName,description,webUrl,createdDateTime,lastModifiedDateTime,siteCollection' },
    scopes: ['Sites.Read.All']
  });
}

export async function listSharePointColumns(config: AppConfig, siteId: string, listId: string): Promise<unknown> {
  return graphRequest(config, 'GET', `/sites/${segment(siteId)}/lists/${segment(listId)}/columns`, {
    scopes: ['Sites.Read.All']
  });
}

export async function getSharePointListItemDelta(
  config: AppConfig,
  siteId: string,
  listId: string,
  token?: string,
  top?: number
): Promise<unknown> {
  return graphRequest(config, 'GET', `/sites/${segment(siteId)}/lists/${segment(listId)}/items/delta`, {
    query: { token, $top: clampTop(top, 50, 100), $expand: 'fields' },
    scopes: ['Sites.Read.All']
  });
}

export async function createSharePointList(
  config: AppConfig,
  siteId: string,
  displayName: string,
  description?: string,
  template = 'genericList',
  columns?: Array<Record<string, unknown>>
): Promise<unknown> {
  return graphRequest(config, 'POST', `/sites/${segment(siteId)}/lists`, {
    body: { displayName, description, list: { template }, columns },
    scopes: ['Sites.Manage.All']
  });
}

export async function updateSharePointList(
  config: AppConfig,
  siteId: string,
  listId: string,
  input: { displayName?: string; description?: string }
): Promise<unknown> {
  return graphRequest(config, 'PATCH', `/sites/${segment(siteId)}/lists/${segment(listId)}`, {
    body: input,
    scopes: ['Sites.Manage.All']
  });
}

export async function deleteSharePointList(config: AppConfig, siteId: string, listId: string): Promise<unknown> {
  return graphRequest(config, 'DELETE', `/sites/${segment(siteId)}/lists/${segment(listId)}`, {
    scopes: ['Sites.Manage.All']
  });
}

export async function createSharePointColumn(
  config: AppConfig,
  siteId: string,
  listId: string,
  definition: Record<string, unknown>
): Promise<unknown> {
  return graphRequest(config, 'POST', `/sites/${segment(siteId)}/lists/${segment(listId)}/columns`, {
    body: definition,
    scopes: ['Sites.Manage.All']
  });
}

export async function updateSharePointColumn(
  config: AppConfig,
  siteId: string,
  listId: string,
  columnId: string,
  definition: Record<string, unknown>
): Promise<unknown> {
  return graphRequest(
    config,
    'PATCH',
    `/sites/${segment(siteId)}/lists/${segment(listId)}/columns/${segment(columnId)}`,
    { body: definition, scopes: ['Sites.Manage.All'] }
  );
}

export async function deleteSharePointColumn(
  config: AppConfig,
  siteId: string,
  listId: string,
  columnId: string
): Promise<unknown> {
  return graphRequest(
    config,
    'DELETE',
    `/sites/${segment(siteId)}/lists/${segment(listId)}/columns/${segment(columnId)}`,
    { scopes: ['Sites.Manage.All'] }
  );
}

function sharePointDriveItemPath(driveId: string, itemId: string): string {
  return `/drives/${segment(driveId)}/items/${segment(itemId)}`;
}

export async function downloadSharePointFile(config: AppConfig, driveId: string, itemId: string): Promise<unknown> {
  return graphRequest(config, 'GET', `${sharePointDriveItemPath(driveId, itemId)}/content`, {
    scopes: ['Files.Read.All'],
    responseType: 'base64',
    maxResponseBytes: 10 * 1024 * 1024
  });
}

export async function renameSharePointDriveItem(
  config: AppConfig,
  driveId: string,
  itemId: string,
  newName: string
): Promise<unknown> {
  return graphRequest(config, 'PATCH', sharePointDriveItemPath(driveId, itemId), {
    body: { name: newName },
    scopes: ['Files.ReadWrite.All']
  });
}

export async function moveSharePointDriveItem(
  config: AppConfig,
  driveId: string,
  itemId: string,
  newParentItemId: string
): Promise<unknown> {
  return graphRequest(config, 'PATCH', sharePointDriveItemPath(driveId, itemId), {
    body: { parentReference: { id: newParentItemId } },
    scopes: ['Files.ReadWrite.All']
  });
}

export async function copySharePointDriveItem(
  config: AppConfig,
  driveId: string,
  itemId: string,
  parentItemId: string,
  newName?: string
): Promise<unknown> {
  return graphRequest(config, 'POST', `${sharePointDriveItemPath(driveId, itemId)}/copy`, {
    body: { parentReference: { driveId, id: parentItemId }, name: newName },
    scopes: ['Files.ReadWrite.All']
  });
}

export async function createSharePointDriveItemLink(
  config: AppConfig,
  driveId: string,
  itemId: string,
  type: 'view' | 'edit' | 'embed' = 'view',
  scope: 'anonymous' | 'organization' | 'users' = 'organization'
): Promise<unknown> {
  return graphRequest(config, 'POST', `${sharePointDriveItemPath(driveId, itemId)}/createLink`, {
    body: { type, scope },
    scopes: ['Files.ReadWrite.All']
  });
}

export async function listSharePointDriveItemPermissions(
  config: AppConfig,
  driveId: string,
  itemId: string
): Promise<unknown> {
  return graphRequest(config, 'GET', `${sharePointDriveItemPath(driveId, itemId)}/permissions`, {
    scopes: ['Files.Read.All']
  });
}

export async function inviteSharePointDriveItem(
  config: AppConfig,
  driveId: string,
  itemId: string,
  recipientsInput: RecipientInput[],
  roles: Array<'read' | 'write'>,
  message?: string,
  requireSignIn = true,
  sendInvitation = true
): Promise<unknown> {
  return graphRequest(config, 'POST', `${sharePointDriveItemPath(driveId, itemId)}/invite`, {
    body: {
      recipients: recipientsInput.map(recipient),
      roles,
      message,
      requireSignIn,
      sendInvitation
    },
    scopes: ['Files.ReadWrite.All']
  });
}

export async function listSharePointDriveVersions(
  config: AppConfig,
  driveId: string,
  itemId: string
): Promise<unknown> {
  return graphRequest(config, 'GET', `${sharePointDriveItemPath(driveId, itemId)}/versions`, {
    scopes: ['Files.Read.All']
  });
}

export async function restoreSharePointDriveVersion(
  config: AppConfig,
  driveId: string,
  itemId: string,
  versionId: string
): Promise<unknown> {
  return graphRequest(
    config,
    'POST',
    `${sharePointDriveItemPath(driveId, itemId)}/versions/${segment(versionId)}/restoreVersion`,
    { scopes: ['Files.ReadWrite.All'] }
  );
}

// Teams additions

export async function getTeam(config: AppConfig, teamId: string): Promise<unknown> {
  return graphRequest(config, 'GET', `/teams/${segment(teamId)}`, {
    scopes: ['Team.ReadBasic.All']
  });
}

export async function getTeamChannel(config: AppConfig, teamId: string, channelId: string): Promise<unknown> {
  return graphRequest(config, 'GET', `/teams/${segment(teamId)}/channels/${segment(channelId)}`, {
    scopes: ['Channel.ReadBasic.All']
  });
}

export async function listTeamMembers(config: AppConfig, teamId: string, top?: number): Promise<unknown> {
  return graphRequest(config, 'GET', `/teams/${segment(teamId)}/members`, {
    query: { $top: clampTop(top, 100, 999) },
    scopes: ['TeamMember.Read.All']
  });
}

export async function listChannelMembers(
  config: AppConfig,
  teamId: string,
  channelId: string,
  top?: number
): Promise<unknown> {
  return graphRequest(config, 'GET', `/teams/${segment(teamId)}/channels/${segment(channelId)}/members`, {
    query: { $top: clampTop(top, 100, 999) },
    scopes: ['ChannelMember.Read.All']
  });
}

export async function createTeamChannel(
  config: AppConfig,
  teamId: string,
  displayName: string,
  description?: string,
  membershipType: 'standard' | 'private' = 'standard',
  members?: ConversationMemberInput[]
): Promise<unknown> {
  return graphRequest(config, 'POST', `/teams/${segment(teamId)}/channels`, {
    body: {
      displayName,
      description,
      membershipType,
      members: membershipType === 'private' ? members?.map(member => conversationMember(config, member)) : undefined
    },
    scopes: ['Channel.Create']
  });
}

export async function updateTeamChannel(
  config: AppConfig,
  teamId: string,
  channelId: string,
  input: { displayName?: string; description?: string }
): Promise<unknown> {
  return graphRequest(config, 'PATCH', `/teams/${segment(teamId)}/channels/${segment(channelId)}`, {
    body: input,
    scopes: ['ChannelSettings.ReadWrite.All']
  });
}

export async function deleteTeamChannel(config: AppConfig, teamId: string, channelId: string): Promise<unknown> {
  return graphRequest(config, 'DELETE', `/teams/${segment(teamId)}/channels/${segment(channelId)}`, {
    scopes: ['Channel.Delete.All']
  });
}

export async function createTeamsChat(
  config: AppConfig,
  chatType: 'oneOnOne' | 'group',
  members: ConversationMemberInput[],
  topic?: string
): Promise<unknown> {
  return graphRequest(config, 'POST', '/chats', {
    body: {
      chatType,
      topic: chatType === 'group' ? topic : undefined,
      members: members.map(member => conversationMember(config, member, 'owner'))
    },
    scopes: ['Chat.Create']
  });
}

export async function listChatMembers(config: AppConfig, chatId: string): Promise<unknown> {
  return graphRequest(config, 'GET', `/chats/${segment(chatId)}/members`, {
    scopes: ['Chat.Read']
  });
}

export async function addChatMember(
  config: AppConfig,
  chatId: string,
  member: ConversationMemberInput,
  visibleHistoryStartDateTime?: string
): Promise<unknown> {
  return graphRequest(config, 'POST', `/chats/${segment(chatId)}/members`, {
    body: { ...conversationMember(config, member, 'owner'), visibleHistoryStartDateTime },
    scopes: ['ChatMember.ReadWrite']
  });
}

export async function removeChatMember(config: AppConfig, chatId: string, membershipId: string): Promise<unknown> {
  return graphRequest(config, 'DELETE', `/chats/${segment(chatId)}/members/${segment(membershipId)}`, {
    scopes: ['ChatMember.ReadWrite']
  });
}

export async function addTeamMember(
  config: AppConfig,
  teamId: string,
  member: ConversationMemberInput
): Promise<unknown> {
  return graphRequest(config, 'POST', `/teams/${segment(teamId)}/members`, {
    body: conversationMember(config, member),
    scopes: ['TeamMember.ReadWrite.All']
  });
}

export async function removeTeamMember(config: AppConfig, teamId: string, membershipId: string): Promise<unknown> {
  return graphRequest(config, 'DELETE', `/teams/${segment(teamId)}/members/${segment(membershipId)}`, {
    scopes: ['TeamMember.ReadWrite.All']
  });
}

export async function addChannelMember(
  config: AppConfig,
  teamId: string,
  channelId: string,
  member: ConversationMemberInput
): Promise<unknown> {
  return graphRequest(config, 'POST', `/teams/${segment(teamId)}/channels/${segment(channelId)}/members`, {
    body: conversationMember(config, member),
    scopes: ['ChannelMember.ReadWrite.All']
  });
}

export async function removeChannelMember(
  config: AppConfig,
  teamId: string,
  channelId: string,
  membershipId: string
): Promise<unknown> {
  return graphRequest(
    config,
    'DELETE',
    `/teams/${segment(teamId)}/channels/${segment(channelId)}/members/${segment(membershipId)}`,
    { scopes: ['ChannelMember.ReadWrite.All'] }
  );
}

export async function softDeleteChannelMessage(
  config: AppConfig,
  teamId: string,
  channelId: string,
  messageId: string,
  replyId?: string
): Promise<unknown> {
  const base = `/teams/${segment(teamId)}/channels/${segment(channelId)}/messages/${segment(messageId)}`;
  const path = replyId ? `${base}/replies/${segment(replyId)}/softDelete` : `${base}/softDelete`;
  return graphRequest(config, 'POST', path, { scopes: ['ChannelMessage.ReadWrite'] });
}

export async function softDeleteChatMessage(config: AppConfig, chatId: string, messageId: string): Promise<unknown> {
  return graphRequest(config, 'POST', `/chats/${segment(chatId)}/messages/${segment(messageId)}/softDelete`, {
    scopes: ['Chat.ReadWrite']
  });
}

// Personal contacts additions

export async function getContact(config: AppConfig, contactId: string): Promise<unknown> {
  return graphRequest(config, 'GET', `/me/contacts/${segment(contactId)}`, {
    scopes: ['Contacts.Read']
  });
}

export async function searchContacts(config: AppConfig, query: string, top?: number): Promise<unknown> {
  const escaped = escapeODataString(query.trim());
  return graphRequest(config, 'GET', '/me/contacts', {
    query: {
      $top: clampTop(top, 20, 100),
      $filter: `startswith(displayName,'${escaped}') or startswith(givenName,'${escaped}') or startswith(surname,'${escaped}')`,
      $select: 'id,displayName,givenName,surname,emailAddresses,businessPhones,mobilePhone,companyName,jobTitle'
    },
    scopes: ['Contacts.Read']
  });
}

export async function listContactFolders(config: AppConfig, top?: number): Promise<unknown> {
  return graphRequest(config, 'GET', '/me/contactFolders', {
    query: { $top: clampTop(top, 25, 100), $select: 'id,displayName,parentFolderId' },
    scopes: ['Contacts.Read']
  });
}

export async function listFolderContacts(config: AppConfig, folderId: string, top?: number): Promise<unknown> {
  return graphRequest(config, 'GET', `/me/contactFolders/${segment(folderId)}/contacts`, {
    query: {
      $top: clampTop(top, 25, 100),
      $select: 'id,displayName,givenName,surname,emailAddresses,businessPhones,mobilePhone,companyName,jobTitle'
    },
    scopes: ['Contacts.Read']
  });
}

export async function createContactFolder(config: AppConfig, displayName: string): Promise<unknown> {
  return graphRequest(config, 'POST', '/me/contactFolders', {
    body: { displayName },
    scopes: ['Contacts.ReadWrite']
  });
}

export async function updateContactFolder(
  config: AppConfig,
  folderId: string,
  displayName: string
): Promise<unknown> {
  return graphRequest(config, 'PATCH', `/me/contactFolders/${segment(folderId)}`, {
    body: { displayName },
    scopes: ['Contacts.ReadWrite']
  });
}

export async function deleteContactFolder(config: AppConfig, folderId: string): Promise<unknown> {
  return graphRequest(config, 'DELETE', `/me/contactFolders/${segment(folderId)}`, {
    scopes: ['Contacts.ReadWrite']
  });
}

// Microsoft Search additions

export async function searchMicrosoft365(
  config: AppConfig,
  entityType: 'message' | 'event' | 'driveItem' | 'listItem' | 'site' | 'chatMessage',
  query: string,
  scopes: string[],
  from = 0,
  size = 25,
  fields?: string[]
): Promise<unknown> {
  return graphRequest(config, 'POST', '/search/query', {
    body: {
      requests: [
        {
          entityTypes: [entityType],
          query: { queryString: query },
          from: Math.max(0, Math.trunc(from)),
          size: clampTop(size, 25, 100),
          fields: fields?.length ? fields : undefined
        }
      ]
    },
    scopes
  });
}
