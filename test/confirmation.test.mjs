import assert from 'node:assert/strict';
import test from 'node:test';

import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { ElicitRequestSchema } from '@modelcontextprotocol/sdk/types.js';
import { z } from 'zod/v4';
import { BrowserConfirmationStore, browserConfirmationStore } from '../dist/browser-confirmation.js';
import { loadConfig } from '../dist/config.js';
import { createGraphMcpServer } from '../dist/mcp-server.js';
import { runWithRequestContext } from '../dist/request-context.js';
import { confirmModule } from '../dist/tools/confirm.js';
import { ToolRegistry } from '../dist/tools/registry.js';
import { jsonResult } from '../dist/tools/results.js';
import { instrumentMcpServer, ToolRuntime } from '../dist/tools/runtime.js';

async function connect(server, elicitationHandler, capabilities = { elicitation: { form: {} } }) {
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  const client = new Client(
    { name: 'native-confirmation-test-client', version: '1.0.0' },
    { capabilities }
  );
  if (elicitationHandler) {
    client.setRequestHandler(ElicitRequestSchema, elicitationHandler);
  }
  await server.connect(serverTransport);
  await client.connect(clientTransport);
  return {
    client,
    async close() {
      await client.close();
      await server.close();
    }
  };
}

function testServer(confirmOperations = ['delete_test_item']) {
  const config = {
    ...loadConfig(),
    transport: 'http',
    inboundAuthDisabled: true,
    roleBasedFiltering: false,
    auditLogEnabled: false,
    confirmOperations
  };
  const registry = new ToolRegistry();
  registry.register({
    name: 'test_delete',
    category: 'test',
    module: 'Test',
    description: 'Delete a test item.',
    isWriteOperation: true,
    requiresConfirmation: true,
    operationType: 'delete_test_item',
    graphScopes: []
  });
  registry.register({
    name: 'confirm_execute',
    category: 'gateway',
    module: 'Operation Confirmation',
    description: 'Execute an approved operation.',
    isWriteOperation: false,
    requiresConfirmation: false,
    graphScopes: []
  });
  const server = new McpServer({ name: 'confirmation-test', version: '1.0.0' });
  instrumentMcpServer(server, new ToolRuntime(config, registry, server));
  confirmModule.register(server, config);
  return { server, config };
}

test('mail send displays a readable client-native preview and decline performs no Graph write', async (t) => {
  browserConfirmationStore.clear();
  t.after(() => browserConfirmationStore.clear());
  const config = {
    ...loadConfig(),
    transport: 'http',
    inboundAuthDisabled: true,
    roleBasedFiltering: false,
    auditLogEnabled: false,
    confirmOperations: ['send_email']
  };
  const server = createGraphMcpServer(config);
  let request;
  const connection = await connect(server, async (incoming) => {
    request = incoming;
    return { action: 'decline' };
  });
  t.after(() => connection.close());

  const tools = await connection.client.listTools();
  assert.equal(tools.tools.some(tool => tool.name === 'confirm_execute'), true);
  assert.equal(tools.tools.some(tool => tool.name === 'mail_send_confirmed'), false);

  const result = await connection.client.callTool({
    name: 'mail_send',
    arguments: {
      subject: 'Client confirmation test',
      body: 'This body must be shown as readable text and must never be sent.',
      to: [{ name: 'Test User', email: 'recipient@example.cn' }]
    }
  });

  assert.equal(request.method, 'elicitation/create');
  assert.equal(request.params.mode, 'form');
  assert.match(request.params.message, /Send Mail/);
  assert.match(request.params.message, /Client confirmation test/);
  assert.match(request.params.message, /Test User <recipient@example\.cn>/);
  assert.match(request.params.message, /This body must be shown as readable text/);
  assert.doesNotMatch(request.params.message, /\{"toolName"/);
  assert.equal(request.params.requestedSchema.properties.confirm.default, false);
  assert.equal(result.structuredContent.executed, false);
  assert.equal(result.structuredContent.confirmationStatus, 'declined');
  assert.match(result.content[0].text, /未执行任何写入/);
  assert.equal(browserConfirmationStore.size(), 0);
});

test('a write executes exactly once only after explicit native confirmation', async (t) => {
  const { server } = testServer();
  let executions = 0;
  server.registerTool(
    'test_delete',
    {
      title: 'Delete Test Item',
      description: 'Delete a test item.',
      inputSchema: { itemId: z.string(), reason: z.string() }
    },
    async () => {
      executions += 1;
      return jsonResult({ deleted: true });
    }
  );
  const connection = await connect(server, async () => ({
    action: 'accept',
    content: { confirm: true }
  }));
  t.after(() => connection.close());

  const result = await connection.client.callTool({
    name: 'test_delete',
    arguments: { itemId: 'item-1', reason: 'test cleanup' }
  });
  assert.equal(result.structuredContent.deleted, true);
  assert.equal(executions, 1);
});

test('accepting the dialog without checking confirm does not execute', async (t) => {
  const { server } = testServer();
  let executions = 0;
  server.registerTool(
    'test_delete',
    { description: 'Delete a test item.', inputSchema: { itemId: z.string() } },
    async () => {
      executions += 1;
      return jsonResult({ deleted: true });
    }
  );
  const connection = await connect(server, async () => ({
    action: 'accept',
    content: { confirm: false }
  }));
  t.after(() => connection.close());

  const result = await connection.client.callTool({
    name: 'test_delete',
    arguments: { itemId: 'item-1' }
  });
  assert.equal(result.structuredContent.executed, false);
  assert.equal(result.structuredContent.confirmationStatus, 'not_confirmed');
  assert.equal(executions, 0);
});

test('clients without form elicitation receive a browser preview and execute only after approval', async (t) => {
  browserConfirmationStore.clear();
  t.after(() => browserConfirmationStore.clear());
  const { server } = testServer();
  let executions = 0;
  server.registerTool(
    'test_delete',
    { description: 'Delete a test item.', inputSchema: { itemId: z.string() } },
    async () => {
      executions += 1;
      return jsonResult({ deleted: true });
    }
  );
  const connection = await connect(server, undefined, {});
  t.after(() => connection.close());

  const result = await connection.client.callTool({
    name: 'test_delete',
    arguments: { itemId: 'item-1' }
  });
  assert.equal(result.isError, undefined);
  assert.equal(result.structuredContent.confirmationStatus, 'awaiting_human_approval');
  assert.match(result.structuredContent.approvalUrl, /\/confirm\/confirm_/);
  assert.equal(result.structuredContent.previewDetails.fields[0].value, 'item-1');
  assert.match(result.content[0].text, /不要改用草稿或其他写工具/);
  assert.equal(executions, 0);

  const token = result.structuredContent.confirmToken;
  const premature = await connection.client.callTool({
    name: 'confirm_execute',
    arguments: { confirmToken: token }
  });
  assert.equal(premature.isError, true);
  assert.equal(premature.structuredContent.error.code, 'human_confirmation_required');
  assert.equal(executions, 0);

  const preparation = browserConfirmationStore.prepareBrowserApproval(token);
  assert.ok(preparation?.browserNonce);
  assert.equal(
    browserConfirmationStore.approveFromBrowser(token, preparation.browserNonce)?.approved,
    true
  );
  const confirmed = await connection.client.callTool({
    name: 'confirm_execute',
    arguments: { confirmToken: token }
  });
  assert.equal(confirmed.structuredContent.deleted, true);
  assert.equal(executions, 1);

  const replay = await connection.client.callTool({
    name: 'confirm_execute',
    arguments: { confirmToken: token }
  });
  assert.equal(replay.isError, true);
  assert.equal(replay.structuredContent.error.code, 'invalid_confirm_token');
  assert.equal(executions, 1);
});

test('a native elicitation transport error falls back to browser confirmation', async (t) => {
  browserConfirmationStore.clear();
  t.after(() => browserConfirmationStore.clear());
  const { server } = testServer();
  let executions = 0;
  server.registerTool(
    'test_delete',
    { description: 'Delete a test item.', inputSchema: { itemId: z.string() } },
    async () => {
      executions += 1;
      return jsonResult({ deleted: true });
    }
  );
  const connection = await connect(server, async () => {
    throw new Error('Client UI cannot render forms');
  });
  t.after(() => connection.close());

  const result = await connection.client.callTool({
    name: 'test_delete',
    arguments: { itemId: 'item-2' }
  });
  assert.equal(result.structuredContent.confirmationStatus, 'awaiting_human_approval');
  assert.match(result.structuredContent.approvalUrl, /\/confirm\/confirm_/);
  assert.equal(executions, 0);
});

test('legacy empty elicitation capability is treated as form support', async (t) => {
  const { server } = testServer();
  let executions = 0;
  server.registerTool(
    'test_delete',
    { description: 'Delete a test item.', inputSchema: { itemId: z.string() } },
    async () => {
      executions += 1;
      return jsonResult({ deleted: true });
    }
  );
  const connection = await connect(
    server,
    async () => ({ action: 'cancel' }),
    { elicitation: {} }
  );
  t.after(() => connection.close());

  const result = await connection.client.callTool({
    name: 'test_delete',
    arguments: { itemId: 'item-1' }
  });
  assert.equal(result.structuredContent.confirmationStatus, 'cancelled');
  assert.equal(executions, 0);
});

test('all confirmation policy gates every confirmable write operation', async (t) => {
  const { server } = testServer(['all']);
  let executions = 0;
  server.registerTool(
    'test_delete',
    { description: 'Delete a test item.', inputSchema: { itemId: z.string() } },
    async () => {
      executions += 1;
      return jsonResult({ deleted: true });
    }
  );
  const connection = await connect(server, async () => ({ action: 'decline' }));
  t.after(() => connection.close());

  const result = await connection.client.callTool({
    name: 'test_delete',
    arguments: { itemId: 'item-1' }
  });
  assert.equal(result.structuredContent.confirmationStatus, 'declined');
  assert.equal(executions, 0);
});

test('automatic mode sends immediately but still gates destructive writes', async (t) => {
  const config = {
    ...loadConfig(),
    transport: 'http',
    inboundAuthDisabled: true,
    roleBasedFiltering: false,
    auditLogEnabled: false,
    sendMode: 'automatic',
    confirmOperations: ['all']
  };
  const registry = new ToolRegistry();
  registry.register({
    name: 'test_send',
    category: 'test',
    module: 'Test',
    description: 'Send a test message.',
    isWriteOperation: true,
    requiresConfirmation: true,
    operationType: 'send_email',
    graphScopes: []
  });
  registry.register({
    name: 'test_delete',
    category: 'test',
    module: 'Test',
    description: 'Delete a test item.',
    isWriteOperation: true,
    requiresConfirmation: true,
    operationType: 'delete_test_item',
    graphScopes: []
  });
  const server = new McpServer({ name: 'automatic-send-test', version: '1.0.0' });
  instrumentMcpServer(server, new ToolRuntime(config, registry, server));
  let sends = 0;
  let deletes = 0;
  let confirmationRequests = 0;
  server.registerTool(
    'test_send',
    { description: 'Send a test message.', inputSchema: { body: z.string() } },
    async () => {
      sends += 1;
      return jsonResult({ sent: true });
    }
  );
  server.registerTool(
    'test_delete',
    { description: 'Delete a test item.', inputSchema: { itemId: z.string() } },
    async () => {
      deletes += 1;
      return jsonResult({ deleted: true });
    }
  );
  const connection = await connect(server, async () => {
    confirmationRequests += 1;
    return { action: 'decline' };
  });
  t.after(() => connection.close());

  const sent = await connection.client.callTool({
    name: 'test_send',
    arguments: { body: 'approved automation' }
  });
  assert.equal(sent.structuredContent.sent, true);
  assert.equal(sends, 1);
  assert.equal(confirmationRequests, 0);

  const deleted = await connection.client.callTool({
    name: 'test_delete',
    arguments: { itemId: 'item-automatic-test' }
  });
  assert.equal(deleted.structuredContent.confirmationStatus, 'declined');
  assert.equal(deletes, 0);
  assert.equal(confirmationRequests, 1);
});

test('browser confirmation tokens are user-bound and expire', async () => {
  let now = 1_000;
  const store = new BrowserConfirmationStore(() => now);
  const entry = {
    name: 'test_delete',
    title: 'Delete Test Item',
    category: 'test',
    module: 'Test',
    description: 'Delete a test item.',
    isWriteOperation: true,
    requiresConfirmation: true,
    operationType: 'delete_test_item',
    graphScopes: []
  };
  let executions = 0;
  const pending = store.create(entry, { itemId: 'item-3' }, 'user-a', 1, async () => {
    executions += 1;
    return jsonResult({ deleted: true });
  });
  const preparation = store.prepareBrowserApproval(pending.confirmToken);
  assert.ok(preparation?.browserNonce);
  store.approveFromBrowser(pending.confirmToken, preparation.browserNonce);

  const wrongUser = await store.execute(pending.confirmToken, 'user-b');
  assert.equal(wrongUser.structuredContent.error.code, 'confirm_user_mismatch');
  assert.equal(executions, 0);

  now += 1_001;
  const expired = await store.execute(pending.confirmToken, 'user-a');
  assert.equal(expired.structuredContent.error.code, 'expired_confirm_token');
  assert.equal(executions, 0);
});

test('web approval does not bypass a role revoked before execution', async () => {
  browserConfirmationStore.clear();
  const config = {
    ...loadConfig(),
    transport: 'http',
    inboundAuthDisabled: false,
    roleBasedFiltering: true,
    auditLogEnabled: false,
    confirmOperations: ['delete_test_item'],
    confirmTtlSeconds: 300
  };
  const registry = new ToolRegistry();
  registry.register({
    name: 'test_delete',
    category: 'test',
    module: 'Test',
    description: 'Delete a test item.',
    isWriteOperation: true,
    requiresConfirmation: true,
    operationType: 'delete_test_item',
    graphScopes: [],
    requiredRoles: ['mcp.test']
  });
  registry.register({
    name: 'confirm_execute',
    category: 'gateway',
    module: 'Operation Confirmation',
    description: 'Execute an approved operation.',
    isWriteOperation: false,
    requiresConfirmation: false,
    graphScopes: []
  });
  const server = new McpServer({ name: 'role-recheck-test', version: '1.0.0' });
  instrumentMcpServer(server, new ToolRuntime(config, registry, server), { userRoles: ['mcp.test'] });
  let executions = 0;
  server.registerTool(
    'test_delete',
    { description: 'Delete a test item.', inputSchema: { itemId: z.string() } },
    async () => {
      executions += 1;
      return jsonResult({ deleted: true });
    }
  );
  confirmModule.register(server, config);

  const initial = await runWithRequestContext(
    { requestId: 'role-before', inboundClaims: { oid: 'user-a', roles: ['mcp.test'] } },
    () => server._registeredTools.test_delete.handler({ itemId: 'item-4' }, {})
  );
  const preparation = browserConfirmationStore.prepareBrowserApproval(initial.structuredContent.confirmToken);
  assert.ok(preparation?.browserNonce);
  browserConfirmationStore.approveFromBrowser(initial.structuredContent.confirmToken, preparation.browserNonce);

  const denied = await runWithRequestContext(
    { requestId: 'role-after', inboundClaims: { oid: 'user-a', roles: [] } },
    () =>
      server._registeredTools.confirm_execute.handler(
        { confirmToken: initial.structuredContent.confirmToken },
        {}
      )
  );
  assert.equal(denied.structuredContent.error.code, 'permission_denied');
  assert.equal(executions, 0);
  browserConfirmationStore.clear();
});
