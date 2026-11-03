import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import test from 'node:test';
import { promisify } from 'node:util';

import { loadConfig } from '../dist/config.js';
import { createGraphMcpServer } from '../dist/mcp-server.js';
import { permissionReport } from '../dist/permissions-report.js';
import { runWithRequestContext } from '../dist/request-context.js';
import { createToolRegistry } from '../dist/tools/index.js';

const run = promisify(execFile);

function gatewayConfig(overrides = {}) {
  return {
    ...loadConfig(),
    transport: 'http',
    inboundAuthDisabled: false,
    enableRawGraphGet: false,
    toolCategories: [],
    disabledGraphScopes: [],
    allowedGraphScopes: [],
    readOnly: false,
    roleBasedFiltering: true,
    auditLogEnabled: false,
    sendMode: 'confirm',
    toolExposureMode: 'hybrid',
    ...overrides
  };
}

test('read-only mode removes every write tool, the confirmation tool and their discovery entries', async () => {
  const config = gatewayConfig({ readOnly: true });
  const registry = createToolRegistry(config);
  assert.equal(registry.getWriteOperations().length, 0);
  assert.ok(registry.getByName('mail_list_messages'));
  assert.equal(registry.getByName('mail_send'), undefined);

  const server = createGraphMcpServer(config, { userRoles: ['mcp.admin'] });
  const names = Object.keys(server._registeredTools);
  assert.ok(!names.includes('confirm_execute'));
  assert.ok(names.includes('mail_get_attachment'));
  assert.match(server.server._instructions ?? '', /read-only/);

  const search = await runWithRequestContext(
    { requestId: 'read-only-search', inboundClaims: { sub: 'admin', roles: ['mcp.admin'] } },
    () => server._registeredTools.gateway_search_tools.handler({ query: '发送邮件', limit: 10, schemaCount: 0 }, {})
  );
  assert.equal(search.structuredContent.tools.some(tool => tool.name === 'mail_send'), false);
});

test('an allow-list hides tools whose scopes it does not cover', () => {
  const registry = createToolRegistry(gatewayConfig({ allowedGraphScopes: ['User.Read', 'Mail.ReadWrite'] }));
  const names = registry.getAll().map(entry => entry.name);
  assert.ok(names.includes('graph_get_me'));
  assert.ok(names.includes('mail_list_messages'), 'Mail.ReadWrite covers Mail.Read');
  assert.ok(names.includes('mail_move_message'));
  assert.ok(!names.includes('mail_send'), 'Mail.Send is not covered');
  assert.ok(!names.includes('calendar_list_view'));
  assert.ok(!names.includes('files_search'));
});

test('the permission report lists required, optional and missing scopes', () => {
  const report = permissionReport(
    gatewayConfig({
      disabledGraphScopes: ['User.Read.All', 'ChannelMessage.Read.All'],
      allowedGraphScopes: []
    })
  );
  assert.ok(report.requiredScopes.includes('Mail.Read'));
  assert.ok(!report.requiredScopes.includes('User.Read.All'));
  const hidden = Object.fromEntries(report.hiddenTools.map(tool => [tool.name, tool]));
  assert.deepEqual(hidden.users_get_profile.missingScopes, ['User.Read.All']);
  assert.deepEqual(hidden.teams_list_channel_messages.missingScopes, ['ChannelMessage.Read.All']);
  assert.deepEqual(
    report.optionalScopes.find(item => item.scope === 'User.Read.All'),
    { scope: 'User.Read.All', enabled: false, tools: ['users_list_memberships'] }
  );

  const readOnly = permissionReport(gatewayConfig({ readOnly: true, allowedGraphScopes: ['Mail.Read', 'Notes.Read'] }));
  assert.deepEqual(readOnly.requiredScopes, ['Mail.Read']);
  assert.equal(readOnly.tools.every(tool => tool.access === 'read'), true);
  assert.ok(readOnly.hiddenTools.some(tool => tool.name === 'mail_send' && tool.reason === 'read_only'));
  assert.deepEqual(readOnly.allowedScopesNotUsed, ['Notes.Read']);
});

test('--list-permissions prints a JSON report without starting the server', async () => {
  const { stdout } = await run(process.execPath, ['dist/index.js', '--list-permissions', '--json'], {
    env: { ...process.env, MCP_TRANSPORT: 'http', MCP_DISABLED_GRAPH_SCOPES: 'User.Read.All' },
    timeout: 10_000
  });
  const report = JSON.parse(stdout);
  assert.equal(report.readOnly, false);
  assert.ok(report.requiredScopes.includes('Chat.Read'));
  assert.ok(report.hiddenTools.some(tool => tool.name === 'users_get_manager'));
});
