import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import test from 'node:test';

test('stdio transport writes only JSON-RPC to stdout', async (t) => {
  const child = spawn(process.execPath, ['dist/index.js', '--stdio'], {
    env: { ...process.env, MCP_TRANSPORT: 'stdio' },
    stdio: ['pipe', 'pipe', 'pipe']
  });
  t.after(() => {
    if (child.exitCode === null) child.kill('SIGTERM');
  });
  let stdout = '';
  const firstLine = new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error(`No stdout line within 10s. Got: ${JSON.stringify(stdout)}`)), 10_000);
    child.stdout.on('data', chunk => {
      stdout += chunk.toString();
      const newline = stdout.indexOf('\n');
      if (newline >= 0) {
        clearTimeout(timer);
        resolve(stdout.slice(0, newline));
      }
    });
    child.once('exit', code => reject(new Error(`stdio server exited early with ${code}`)));
  });
  child.stdin.write(
    `${JSON.stringify({
      jsonrpc: '2.0',
      id: 1,
      method: 'initialize',
      params: { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'stdio-test', version: '1.0.0' } }
    })}\n`
  );

  const message = JSON.parse(await firstLine);
  assert.equal(message.id, 1);
  assert.equal(message.result.serverInfo.name, 'ms-365-21v-mcp-server');
});
