import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import net from 'node:net';
import test from 'node:test';

import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js';
import { ElicitRequestSchema } from '@modelcontextprotocol/sdk/types.js';

async function availablePort() {
  const server = net.createServer();
  await new Promise((resolve, reject) => {
    server.once('error', reject);
    server.listen(0, '127.0.0.1', resolve);
  });
  const address = server.address();
  const port = typeof address === 'object' && address ? address.port : 0;
  await new Promise(resolve => server.close(resolve));
  return port;
}

async function waitForHealth(url, child, output) {
  const deadline = Date.now() + 10_000;
  while (Date.now() < deadline) {
    if (child.exitCode !== null) {
      throw new Error(`HTTP test server exited early (${child.exitCode}).\n${output()}`);
    }
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {
      // Server is still starting.
    }
    await new Promise(resolve => setTimeout(resolve, 50));
  }
  throw new Error(`Timed out waiting for HTTP test server.\n${output()}`);
}

test('Streamable HTTP supports native elicitation and a readable browser fallback', async (t) => {
  const port = await availablePort();
  let output = '';
  const child = spawn(process.execPath, ['dist/index.js', '--http'], {
    cwd: process.cwd(),
    env: {
      ...process.env,
      MCP_HTTP_HOST: '127.0.0.1',
      MCP_HTTP_PORT: String(port),
      MCP_HTTP_PATH: '/mcp',
      MCP_PUBLIC_BASE_URL: `http://127.0.0.1:${port}`,
      MCP_RESOURCE_URL: `http://127.0.0.1:${port}/mcp`,
      MCP_OAUTH_BRIDGE_ISSUER: `http://127.0.0.1:${port}`,
      MCP_OAUTH_HTTP_PATH: '',
      MCP_HTTP_STATELESS: 'false',
      MCP_INBOUND_AUTH_DISABLED: 'true',
      MCP_OAUTH_BRIDGE_ENABLED: 'false',
      MCP_ROLE_BASED_FILTERING: 'false',
      MCP_AUDIT_LOG_ENABLED: 'false',
      MCP_CONFIRM_OPERATIONS: 'send_email'
    },
    stdio: ['ignore', 'pipe', 'pipe']
  });
  child.stdout.on('data', chunk => {
    output += chunk.toString();
  });
  child.stderr.on('data', chunk => {
    output += chunk.toString();
  });
  t.after(() => {
    if (child.exitCode === null) child.kill('SIGTERM');
  });

  await waitForHealth(`http://127.0.0.1:${port}/healthz`, child, () => output);

  const transport = new StreamableHTTPClientTransport(new URL(`http://127.0.0.1:${port}/mcp`));
  const client = new Client(
    { name: 'http-native-confirmation-test', version: '1.0.0' },
    { capabilities: { elicitation: { form: {} } } }
  );
  let elicitationCount = 0;
  client.setRequestHandler(ElicitRequestSchema, async request => {
    elicitationCount += 1;
    assert.match(request.params.message, /HTTP confirmation test/);
    return { action: 'decline' };
  });
  await client.connect(transport);
  t.after(() => client.close());

  const result = await client.callTool({
    name: 'mail_send',
    arguments: {
      subject: 'HTTP confirmation test',
      body: 'This message must not reach Graph.',
      to: [{ email: 'recipient@example.cn' }]
    }
  });

  assert.equal(elicitationCount, 1);
  assert.equal(result.structuredContent.executed, false);
  assert.equal(result.structuredContent.confirmationStatus, 'declined');

  const fallbackTransport = new StreamableHTTPClientTransport(new URL(`http://127.0.0.1:${port}/mcp`));
  const fallbackClient = new Client(
    { name: 'http-browser-confirmation-test', version: '1.0.0' },
    { capabilities: {} }
  );
  await fallbackClient.connect(fallbackTransport);
  t.after(() => fallbackClient.close());
  const fallback = await fallbackClient.callTool({
    name: 'mail_send',
    arguments: {
      subject: 'Readable browser fallback',
      body: 'The browser page must show this exact body.',
      to: [{ name: 'Test User', email: 'recipient@example.cn' }]
    }
  });
  assert.equal(fallback.structuredContent.confirmationStatus, 'awaiting_human_approval');
  const browserPage = await fetch(fallback.structuredContent.approvalUrl);
  assert.equal(browserPage.status, 200);
  const browserHtml = await browserPage.text();
  assert.match(browserHtml, /Readable browser fallback/);
  assert.match(browserHtml, /Test User &lt;recipient@example\.cn&gt;/);
  assert.match(browserHtml, /The browser page must show this exact body/);
  assert.doesNotMatch(browserHtml, /"toolName"/);

  const invalidBrowserRoute = await fetch(`http://127.0.0.1:${port}/confirm/not-a-token`);
  assert.equal(invalidBrowserRoute.status, 410);
});
