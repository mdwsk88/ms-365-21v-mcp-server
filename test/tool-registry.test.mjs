import assert from 'node:assert/strict';
import test from 'node:test';

import { loadConfig } from '../dist/config.js';
import { createGraphMcpServer } from '../dist/mcp-server.js';
import { runWithRequestContext } from '../dist/request-context.js';
import { createToolRegistry } from '../dist/tools/index.js';

function gatewayConfig(overrides = {}) {
  return {
    ...loadConfig(),
    toolExposureMode: 'direct',
    transport: 'http',
    enableRawGraphGet: false,
    toolCategories: [],
    disabledGraphScopes: [],
    roleBasedFiltering: true,
    auditLogEnabled: false,
    confirmOperations: ['production-confirmation-enabled'],
    ...overrides
  };
}

function toolNames(server) {
  return Object.keys(server._registeredTools);
}

test('registry exposes the complete 21V tool catalog and records authorization metadata', () => {
  const registry = createToolRegistry(gatewayConfig());
  assert.equal(registry.getAll().length, 148);
  assert.deepEqual(registry.getAllCategories(), [
    'calendar',
    'contacts',
    'drive',
    'gateway',
    'groups',
    'mail',
    'search',
    'sharepoint',
    'teams',
    'users'
  ]);

  const sendMail = registry.getByName('mail_send');
  assert.equal(sendMail?.requiredRole, 'mcp.mail');
  assert.equal(sendMail?.isWriteOperation, true);
  assert.equal(sendMail?.operationType, 'send_email');
  assert.deepEqual(sendMail?.graphScopes, ['Mail.Send']);
  assert.deepEqual(sendMail?.requiredRoles, ['mcp.mail']);
  assert.equal(registry.getByName('smart_mail_digest'), undefined);
  assert.deepEqual(registry.getByName('search_files')?.requiredRoles, ['mcp.search', 'mcp.drive']);
  assert.equal(registry.getByCategory('groups').length, 5);
  assert.equal(registry.getByCategory('search').length, 5);
  assert.deepEqual(
    registry
      .getAll()
      .filter((entry) => !['gateway', 'debug'].includes(entry.category) && entry.graphScopes.length === 0),
    []
  );
});

test('disabled Graph scopes remove dependent tools from both registry and MCP catalog', () => {
  const disabledGraphScopes = [
    'Channel.Create',
    'Channel.Delete.All',
    'ChannelMember.Read.All',
    'ChannelMember.ReadWrite.All',
    'ChannelMessage.Read.All',
    'ChannelMessage.ReadWrite',
    'ChannelMessage.Send',
    'ChannelSettings.ReadWrite.All',
    'ChatMember.ReadWrite',
    'GroupMember.Read.All',
    'TeamMember.Read.All',
    'TeamMember.ReadWrite.All',
    'User.Read.All'
  ];
  const config = gatewayConfig({ disabledGraphScopes });
  const registry = createToolRegistry(config);
  const server = createGraphMcpServer(config, { bypassRoleFiltering: true });
  const disabled = new Set(disabledGraphScopes.map(scope => scope.toLowerCase()));

  assert.equal(registry.getAll().length, 126);
  assert.equal(toolNames(server).length, 126);
  assert.equal(
    registry.getAll().filter(tool => tool.graphScopes.some(scope => disabled.has(scope.toLowerCase()))).length,
    0
  );
  assert.equal(registry.getByName('users_get_profile'), undefined);
  assert.equal(registry.getByName('groups_list_members'), undefined);
  assert.equal(registry.getByName('teams_send_channel_message'), undefined);
  assert.equal(server._registeredTools.teams_send_channel_message, undefined);

  assert.ok(registry.getByName('users_search'));
  // Least-privileged permissions are already granted for these two, so withholding User.Read.All keeps them.
  assert.deepEqual(registry.getByName('users_list_direct_reports')?.graphScopes, ['User.Read', 'User.ReadBasic.All']);
  assert.deepEqual(registry.getByName('users_list_memberships')?.optionalGraphScopes, ['User.Read.All']);
  assert.ok(registry.getByName('groups_check_my_memberships'));
  assert.ok(registry.getByName('teams_send_chat_message'));
});

test('category filtering keeps base auth tools and only selected business modules', () => {
  const config = gatewayConfig({ toolCategories: ['mail'], roleBasedFiltering: false });
  const server = createGraphMcpServer(config);
  assert.equal(toolNames(server).length, 24);
  assert.ok(toolNames(server).includes('auth_status'));
  assert.ok(toolNames(server).includes('mail_send'));
  assert.ok(!toolNames(server).includes('calendar_list_events'));
});

test('role filtering exposes no business tools without roles and module tools with a matching role', () => {
  const config = gatewayConfig();
  const roleless = createGraphMcpServer(config, { userRoles: [] });
  assert.deepEqual(toolNames(roleless).sort(), ['auth_status', 'confirm_execute']);

  const mailUser = createGraphMcpServer(config, { userRoles: ['mcp.mail'] });
  assert.equal(toolNames(mailUser).length, 24);
  assert.ok(toolNames(mailUser).includes('mail_list_messages'));
  assert.ok(!toolNames(mailUser).includes('calendar_list_events'));

  const admin = createGraphMcpServer(config, { userRoles: ['mcp.admin'] });
  assert.equal(toolNames(admin).length, 148);

  const groupsUser = createGraphMcpServer(config, { userRoles: ['mcp.groups'] });
  assert.equal(toolNames(groupsUser).length, 7);
  assert.ok(toolNames(groupsUser).includes('groups_list_members'));

  const searchOnlyUser = createGraphMcpServer(config, { userRoles: ['mcp.search'] });
  assert.deepEqual(toolNames(searchOnlyUser).sort(), ['auth_status', 'confirm_execute']);

  const fileSearchUser = createGraphMcpServer(config, { userRoles: ['mcp.search', 'mcp.drive'] });
  assert.ok(toolNames(fileSearchUser).includes('search_files'));
  assert.ok(!toolNames(fileSearchUser).includes('search_mail'));
});

test('invocation-level authorization denies a direct call even when the catalog contains the tool', async () => {
  const config = gatewayConfig();
  const catalogServer = createGraphMcpServer(config, { bypassRoleFiltering: true });
  const handler = catalogServer._registeredTools.mail_send.handler;

  const result = await runWithRequestContext(
    {
      requestId: 'role-test',
      userAssertion: 'not-used-because-role-check-runs-first',
      inboundClaims: { sub: 'user-without-mail-role', roles: ['mcp.calendar'] }
    },
    () => handler({}, {})
  );

  assert.equal(result.isError, true);
  assert.match(result.content[0].text, /permission_denied/i);
  assert.match(result.content[0].text, /mcp\.mail/);
});

test('invocation-level search authorization requires both search and source roles', async () => {
  const config = gatewayConfig();
  const catalogServer = createGraphMcpServer(config, { bypassRoleFiltering: true });
  const handler = catalogServer._registeredTools.search_mail.handler;

  const result = await runWithRequestContext(
    {
      requestId: 'search-role-test',
      userAssertion: 'not-used-because-role-check-runs-first',
      inboundClaims: { sub: 'search-only-user', roles: ['mcp.search'] }
    },
    () => handler({ query: 'budget' }, {})
  );

  assert.equal(result.isError, true);
  assert.match(result.content[0].text, /permission_denied/i);
  assert.match(result.content[0].text, /mcp\.search/);
  assert.match(result.content[0].text, /mcp\.mail/);
});
