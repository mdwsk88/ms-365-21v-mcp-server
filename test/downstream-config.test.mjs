import assert from 'node:assert/strict';
import test from 'node:test';

import { parseDownstreamServices } from '../dist/gateway/proxy.js';

test('downstream MCP declarations parse into disabled-forwarding-safe passthrough descriptors', () => {
  assert.deepEqual(parseDownstreamServices(undefined), []);
  assert.deepEqual(parseDownstreamServices('hr:https://hr.internal/mcp,fleet:http://fleet.internal:3002/mcp'), [
    {
      id: 'hr',
      name: 'hr',
      url: 'https://hr.internal/mcp',
      toolPrefix: 'hr_',
      authMode: 'passthrough',
      enabled: true
    },
    {
      id: 'fleet',
      name: 'fleet',
      url: 'http://fleet.internal:3002/mcp',
      toolPrefix: 'fleet_',
      authMode: 'passthrough',
      enabled: true
    }
  ]);
});

test('downstream declarations reject duplicate IDs and non-HTTP URLs', () => {
  assert.throws(
    () => parseDownstreamServices('hr:https://one.internal/mcp,hr:https://two.internal/mcp'),
    /Duplicate downstream MCP id/
  );
  assert.throws(() => parseDownstreamServices('hr:ftp://hr.internal/mcp'), /must use HTTP or HTTPS/);
});
