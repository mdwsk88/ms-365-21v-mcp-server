import assert from 'node:assert/strict';
import test from 'node:test';

import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod/v4';
import { browserConfirmationStore } from '../dist/browser-confirmation.js';
import { loadConfig } from '../dist/config.js';
import { createGraphMcpServer } from '../dist/mcp-server.js';
import { runWithRequestContext } from '../dist/request-context.js';
import { DynamicToolCatalog } from '../dist/tools/discovery.js';
import { createToolRegistry } from '../dist/tools/index.js';
import { ToolRegistry } from '../dist/tools/registry.js';
import { jsonResult } from '../dist/tools/results.js';
import { ToolRuntime } from '../dist/tools/runtime.js';

function config(overrides = {}) {
  return {
    ...loadConfig(),
    transport: 'http',
    inboundAuthDisabled: false,
    roleBasedFiltering: true,
    auditLogEnabled: false,
    disabledGraphScopes: [],
    confirmOperations: ['delete_test_item'],
    ...overrides
  };
}

function toolNames(server) {
  return Object.keys(server._registeredTools).sort();
}

function callAs(server, roles, toolName, parameters) {
  return runWithRequestContext(
    {
      requestId: `discovery-${toolName}`,
      inboundClaims: { sub: 'discovery-user', roles }
    },
    () => server._registeredTools[toolName].handler(parameters, {})
  );
}

test('discovery mode exposes only gateway controls while searching the authorized catalog', async () => {
  const server = createGraphMcpServer(config({ toolExposureMode: 'discovery' }), {
    userRoles: ['mcp.mail']
  });
  assert.deepEqual(toolNames(server), [
    'auth_status',
    'confirm_execute',
    'gateway_execute_tool',
    'gateway_get_tool_schema',
    'gateway_search_tools'
  ]);

  const search = await callAs(server, ['mcp.mail'], 'gateway_search_tools', {
    query: '查看未读邮件',
    limit: 10
  });
  const names = search.structuredContent.tools.map(tool => tool.name);
  assert.ok(names.includes('mail_list_messages'));
  assert.equal(names.some(name => name.startsWith('calendar_')), false);

  const schema = await callAs(server, ['mcp.mail'], 'gateway_get_tool_schema', {
    toolName: 'mail_list_messages'
  });
  assert.equal(schema.structuredContent.name, 'mail_list_messages');
  assert.equal(schema.structuredContent.inputSchema.type, 'object');
  assert.ok(schema.structuredContent.inputSchema.properties.top);

  const unauthorized = await callAs(server, ['mcp.mail'], 'gateway_get_tool_schema', {
    toolName: 'calendar_list_events'
  });
  assert.equal(unauthorized.isError, true);
  assert.equal(unauthorized.structuredContent.error.code, 'tool_not_available');
});

test('hybrid mode keeps the configured fast lane and discovers the remaining tools', async () => {
  const server = createGraphMcpServer(
    config({
      toolExposureMode: 'hybrid',
      directToolCategories: ['smart'],
      directTools: ['mail_list_messages']
    }),
    { userRoles: ['mcp.mail', 'mcp.calendar', 'mcp.smart'] }
  );
  const names = toolNames(server);
  assert.ok(names.includes('mail_list_messages'));
  assert.ok(names.includes('smart_mail_digest'));
  assert.ok(names.includes('smart_calendar_conflicts'));
  assert.ok(names.includes('gateway_search_tools'));
  assert.equal(names.includes('mail_send'), false);
  assert.equal(names.includes('calendar_list_events'), false);

  const search = await callAs(server, ['mcp.mail', 'mcp.calendar', 'mcp.smart'], 'gateway_search_tools', {
    query: '发送邮件',
    category: 'mail'
  });
  const send = search.structuredContent.tools.find(tool => tool.name === 'mail_send');
  assert.ok(send);
  assert.equal(send.direct, false);
});

test('every deployed business tool can publish a discovery JSON schema', async () => {
  const runtimeConfig = config({ toolExposureMode: 'discovery' });
  const server = createGraphMcpServer(runtimeConfig, { userRoles: ['mcp.admin'] });
  const entries = createToolRegistry(runtimeConfig)
    .getAll()
    .filter(entry => entry.category !== 'gateway');
  assert.equal(entries.length, 148);
  for (const entry of entries) {
    const schema = await callAs(server, ['mcp.admin'], 'gateway_get_tool_schema', {
      toolName: entry.name
    });
    assert.equal(schema.isError, undefined, entry.name);
    assert.equal(schema.structuredContent.inputSchema.type, 'object', entry.name);
  }
});

test('dynamic execution validates input and reuses invocation-level role checks', async () => {
  const runtimeConfig = config({ toolExposureMode: 'discovery', sendMode: 'automatic' });
  const registry = new ToolRegistry();
  registry.register({
    name: 'test_echo',
    category: 'test',
    module: 'Test',
    description: 'Echo a value. 中文意图: 回显内容。',
    isWriteOperation: false,
    requiresConfirmation: false,
    graphScopes: [],
    requiredRoles: ['mcp.test']
  });
  const server = new McpServer({ name: 'dynamic-execution-test', version: '1.0.0' });
  const runtime = new ToolRuntime(runtimeConfig, registry, server);
  const catalog = new DynamicToolCatalog(runtimeConfig);
  let executions = 0;
  catalog.capture(
    'test_echo',
    registry.getByName('test_echo'),
    {
      title: 'Echo',
      description: 'Echo a value. 中文意图: 回显内容。',
      inputSchema: { value: z.string().min(1) }
    },
    async ({ value }) => {
      executions += 1;
      return jsonResult({ value });
    }
  );

  const denied = await runWithRequestContext(
    { requestId: 'dynamic-denied', inboundClaims: { sub: 'user-a', roles: ['mcp.mail'] } },
    () => catalog.execute('test_echo', { value: 'hidden' }, {}, runtime)
  );
  assert.equal(denied.isError, true);
  assert.equal(denied.structuredContent.error.code, 'tool_not_available');
  assert.equal(executions, 0);

  const invalid = await runWithRequestContext(
    { requestId: 'dynamic-invalid', inboundClaims: { sub: 'user-a', roles: ['mcp.test'] } },
    () => catalog.execute('test_echo', { value: '' }, {}, runtime)
  );
  assert.equal(invalid.isError, true);
  assert.equal(invalid.structuredContent.error.code, 'invalid_tool_arguments');
  assert.equal(executions, 0);

  const allowed = await runWithRequestContext(
    { requestId: 'dynamic-allowed', inboundClaims: { sub: 'user-a', roles: ['mcp.test'] } },
    () => catalog.execute('test_echo', { value: 'hello' }, {}, runtime)
  );
  assert.equal(allowed.structuredContent.value, 'hello');
  assert.equal(executions, 1);
});

test('dynamic execution supports zero-argument tools with SDK extra as the sole callback argument', async () => {
  const runtimeConfig = config({ toolExposureMode: 'discovery' });
  const registry = new ToolRegistry();
  registry.register({
    name: 'test_no_input',
    category: 'test',
    module: 'Test',
    description: 'Return a fixed value.',
    isWriteOperation: false,
    requiresConfirmation: false,
    graphScopes: [],
    requiredRoles: ['mcp.test']
  });
  const server = new McpServer({ name: 'dynamic-no-input-test', version: '1.0.0' });
  const runtime = new ToolRuntime(runtimeConfig, registry, server);
  const catalog = new DynamicToolCatalog(runtimeConfig);
  let receivedArguments = 0;
  catalog.capture(
    'test_no_input',
    registry.getByName('test_no_input'),
    { title: 'No Input', description: 'Return a fixed value.' },
    async (...args) => {
      receivedArguments = args.length;
      return jsonResult({ ok: true });
    }
  );
  const result = await runWithRequestContext(
    { requestId: 'dynamic-no-input', inboundClaims: { sub: 'user-a', roles: ['mcp.test'] } },
    () => catalog.execute('test_no_input', {}, { transportMetadata: true }, runtime)
  );
  assert.equal(result.structuredContent.ok, true);
  assert.equal(receivedArguments, 1);
});

test('dynamic writes still require the same human confirmation policy', async (t) => {
  browserConfirmationStore.clear();
  t.after(() => browserConfirmationStore.clear());
  const runtimeConfig = config({ toolExposureMode: 'discovery', sendMode: 'confirm' });
  const registry = new ToolRegistry();
  registry.register({
    name: 'test_send',
    category: 'test',
    module: 'Test',
    description: 'Send a test message.',
    isWriteOperation: true,
    requiresConfirmation: true,
    operationType: 'send_email',
    graphScopes: [],
    requiredRoles: ['mcp.test']
  });
  const server = new McpServer({ name: 'dynamic-confirmation-test', version: '1.0.0' });
  const runtime = new ToolRuntime(runtimeConfig, registry, server);
  const catalog = new DynamicToolCatalog(runtimeConfig);
  let executions = 0;
  catalog.capture(
    'test_send',
    registry.getByName('test_send'),
    {
      title: 'Send Test Message',
      description: 'Send a test message.',
      inputSchema: { subject: z.string().min(1), body: z.string().min(1) }
    },
    async () => {
      executions += 1;
      return jsonResult({ sent: true });
    }
  );

  const result = await runWithRequestContext(
    { requestId: 'dynamic-confirm', inboundClaims: { sub: 'user-a', roles: ['mcp.test'] } },
    () => catalog.execute('test_send', { subject: 'Review me', body: 'Not sent yet' }, {}, runtime)
  );
  assert.equal(result.structuredContent.confirmationStatus, 'awaiting_human_approval');
  assert.equal(executions, 0);
  assert.equal(browserConfirmationStore.size(), 1);
});
