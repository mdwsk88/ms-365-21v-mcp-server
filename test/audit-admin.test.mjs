import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';

import { createMcpExpressApp } from '@modelcontextprotocol/sdk/server/express.js';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { mountAdminRoutes } from '../dist/admin-routes.js';
import { AuditLogger } from '../dist/audit-log.js';
import { loadConfig } from '../dist/config.js';
import { gatewayMetrics } from '../dist/gateway/metrics.js';
import { createGraphMcpServer } from '../dist/mcp-server.js';
import { runWithRequestContext } from '../dist/request-context.js';
import { ToolRegistry } from '../dist/tools/registry.js';
import { jsonResult } from '../dist/tools/results.js';
import { instrumentMcpServer, ToolRuntime } from '../dist/tools/runtime.js';

test('audit logger writes redacted JSONL with mode 0600 and summarizes calls', async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'mcp-audit-'));
  const logger = new AuditLogger(true, directory);
  const timestamp = new Date().toISOString();
  await logger.write({
    timestamp,
    requestId: 'request-1',
    userId: 'user-1',
    userDisplayName: 'Test User',
    toolName: 'mail_send',
    toolCategory: 'mail',
    isWriteOperation: true,
    parameters: {
      subject: 'Visible subject',
      body: 'Secret mail body',
      confirmToken: 'confirm-must-not-be-logged',
      access_token: 'access-token-must-not-be-logged',
      clientSecret: 'client-secret-must-not-be-logged',
      nested: { content: 'Secret file content', filename: 'report.txt' }
    },
    duration: 12,
    success: false,
    errorCode: 'permission_denied',
    graphScopes: ['Mail.Send']
  });

  const filePath = logger.filePath(new Date(timestamp));
  const stat = await fs.stat(filePath);
  assert.equal(stat.mode & 0o777, 0o600);
  const entry = JSON.parse((await fs.readFile(filePath, 'utf8')).trim());
  assert.equal(entry.parameters.subject, 'Visible subject');
  assert.equal(entry.parameters.body, '[REDACTED]');
  assert.equal(entry.parameters.confirmToken, '[REDACTED]');
  assert.equal(entry.parameters.access_token, '[REDACTED]');
  assert.equal(entry.parameters.clientSecret, '[REDACTED]');
  assert.equal(entry.parameters.nested.content, '[REDACTED]');
  assert.equal(entry.parameters.nested.filename, 'report.txt');

  const summary = await logger.summary(7);
  assert.equal(summary.totalCalls, 1);
  assert.equal(summary.failedCalls, 1);
  assert.equal(summary.byUser['user-1'].calls, 1);
  assert.equal(summary.byTool.mail_send.errors, 1);
});

test('audit logger recovers after a failed queued write', async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'mcp-audit-recovery-'));
  const blockedPath = path.join(directory, 'blocked');
  await fs.writeFile(blockedPath, 'not a directory');
  const logger = new AuditLogger(true, blockedPath);
  const entry = {
    timestamp: new Date().toISOString(),
    requestId: 'recovery-1',
    userId: 'user-1',
    toolName: 'auth_status',
    toolCategory: 'gateway',
    isWriteOperation: false,
    parameters: {},
    duration: 1,
    success: true
  };

  await assert.rejects(() => logger.write(entry));
  await fs.rm(blockedPath);
  await logger.write({ ...entry, requestId: 'recovery-2' });

  const recovered = JSON.parse((await fs.readFile(logger.filePath(new Date(entry.timestamp)), 'utf8')).trim());
  assert.equal(recovered.requestId, 'recovery-2');
});

test('tool runtime records denied calls in audit and Prometheus metrics', async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'mcp-runtime-audit-'));
  const config = {
    ...loadConfig(),
    transport: 'http',
    toolCategories: [],
    roleBasedFiltering: true,
    auditLogEnabled: true,
    auditLogPath: directory
  };
  gatewayMetrics.resetForTests();
  const server = createGraphMcpServer(config, { bypassRoleFiltering: true });
  const result = await runWithRequestContext(
    {
      requestId: 'denied-request',
      inboundClaims: { oid: 'user-oid', name: 'Denied User', roles: ['mcp.calendar'] }
    },
    () =>
      server._registeredTools.mail_send.handler(
        { subject: 'Subject', body: 'Sensitive', to: [{ email: 'person@example.cn' }] },
        {}
      )
  );
  assert.equal(result.isError, true);

  const logger = new AuditLogger(true, directory);
  const entry = JSON.parse((await fs.readFile(logger.filePath(), 'utf8')).trim());
  assert.equal(entry.requestId, 'denied-request');
  assert.equal(entry.userId, 'user-oid');
  assert.equal(entry.success, false);
  assert.equal(entry.errorCode, 'permission_denied');
  assert.equal(entry.parameters.body, '[REDACTED]');
  assert.match(gatewayMetrics.prometheus(), /mcp_tool_errors_total\{tool="mail_send",category="mail"\} 1/);
});

test('no-input tools never audit MCP transport metadata as business parameters', async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'mcp-no-input-audit-'));
  const config = {
    ...loadConfig(),
    transport: 'http',
    roleBasedFiltering: false,
    auditLogEnabled: true,
    auditLogPath: directory
  };
  const registry = new ToolRegistry();
  registry.register({
    name: 'status_check',
    category: 'gateway',
    module: 'Test',
    description: '',
    isWriteOperation: false,
    requiresConfirmation: false,
    graphScopes: []
  });
  const server = new McpServer({ name: 'audit-signature-test', version: '1.0.0' });
  instrumentMcpServer(server, new ToolRuntime(config, registry, server));
  server.registerTool('status_check', { description: 'Status.' }, async () => jsonResult({ ok: true }));

  await runWithRequestContext({ requestId: 'no-input-request' }, () =>
    server._registeredTools.status_check.handler({
      signal: {},
      requestInfo: { headers: { authorization: 'Bearer must-never-be-a-parameter' } }
    })
  );

  const logger = new AuditLogger(true, directory);
  const entry = JSON.parse((await fs.readFile(logger.filePath(), 'utf8')).trim());
  assert.deepEqual(entry.parameters, {});
});

test('admin REST API requires its static token and mounts under the public prefix', async (t) => {
  const config = {
    ...loadConfig(),
    transport: 'http',
    toolCategories: [],
    auditLogEnabled: false,
    confirmOperations: ['production-confirmation-enabled'],
    adminToken: 'test-admin-token',
    publicPathPrefix: '/tools/MCP_21V'
  };
  const app = createMcpExpressApp({ host: '127.0.0.1' });
  mountAdminRoutes(app, config);
  const server = await new Promise((resolve) => {
    const listening = app.listen(0, '127.0.0.1', () => resolve(listening));
  });
  t.after(() => new Promise((resolve) => server.close(resolve)));
  const address = server.address();
  const baseUrl = `http://127.0.0.1:${address.port}`;

  const unauthorized = await fetch(`${baseUrl}/admin/tools`);
  assert.equal(unauthorized.status, 401);

  const headers = { Authorization: 'Bearer test-admin-token' };
  const toolsResponse = await fetch(`${baseUrl}/tools/MCP_21V/admin/tools`, { headers });
  assert.equal(toolsResponse.status, 200);
  const tools = await toolsResponse.json();
  assert.equal(tools.tools.length, 150);
  const sendMail = tools.tools.find((tool) => tool.name === 'mail_send');
  assert.equal(sendMail.title, 'Send Mail');
  assert.match(sendMail.description, /Send an email as the signed-in user/);

  const categoriesResponse = await fetch(`${baseUrl}/admin/tools/categories`, { headers });
  assert.equal(categoriesResponse.status, 200);
  const categories = await categoriesResponse.json();
  assert.ok(categories.categories.some((value) => value.category === 'sharepoint' && value.toolCount === 33));

  const metricsResponse = await fetch(`${baseUrl}/admin/metrics`, { headers });
  assert.equal(metricsResponse.status, 200);
  assert.match(await metricsResponse.text(), /mcp_gateway_uptime_seconds/);
});
