import assert from 'node:assert/strict';
import test from 'node:test';

import { ConfigError, loadConfig, requireRemoteConfigured } from '../dist/config.js';

const ENV_KEYS = [
  'MCP_OAUTH_BRIDGE_ENABLED',
  'MCP_OAUTH_BRIDGE_MICROSOFT_CLIENT_TYPE',
  'MCP_SEND_MODE',
  'MCP_TOOL_EXPOSURE_MODE',
  'MCP_DIRECT_TOOL_CATEGORIES',
  'MCP_DIRECT_TOOLS',
  'MS_CLIENT_ID',
  'MS_CLIENT_SECRET',
  'MS_OAUTH_CLIENT_ID',
  'MS_OAUTH_CLIENT_SECRET',
  'MS_PUBLIC_CLIENT_ID',
  'MS_TENANT_ID',
  'MCP_TOKEN_AUDIENCE'
];

function withEnvironment(values, callback) {
  const previous = Object.fromEntries(ENV_KEYS.map(key => [key, process.env[key]]));
  try {
    for (const key of ENV_KEYS) delete process.env[key];
    Object.assign(process.env, values);
    return callback();
  } finally {
    for (const key of ENV_KEYS) {
      const value = previous[key];
      if (value === undefined) delete process.env[key];
      else process.env[key] = value;
    }
  }
}

const common = {
  MCP_OAUTH_BRIDGE_ENABLED: 'true',
  MS_TENANT_ID: 'tenant-id',
  MS_CLIENT_ID: 'api-client-id',
  MS_CLIENT_SECRET: 'api-client-secret',
  MCP_TOKEN_AUDIENCE: 'api://api-client-id'
};

test('public mode remains backward compatible with MS_PUBLIC_CLIENT_ID', () => {
  withEnvironment(
    {
      ...common,
      MCP_OAUTH_BRIDGE_MICROSOFT_CLIENT_TYPE: 'public',
      MS_PUBLIC_CLIENT_ID: 'legacy-public-client-id'
    },
    () => {
      const config = loadConfig();
      assert.equal(config.oauthBridgeMicrosoftClientType, 'public');
      assert.equal(config.oauthBridgeMicrosoftClientId, 'legacy-public-client-id');
      assert.equal(config.oauthBridgeMicrosoftClientSecret, undefined);
      assert.doesNotThrow(() => requireRemoteConfigured(config));
    }
  );
});

test('confidential web mode can reuse the MCP API app as one Entra registration', () => {
  withEnvironment(
    {
      ...common,
      MCP_OAUTH_BRIDGE_MICROSOFT_CLIENT_TYPE: 'confidential_web'
    },
    () => {
      const config = loadConfig();
      assert.equal(config.oauthBridgeMicrosoftClientType, 'confidential_web');
      assert.equal(config.oauthBridgeMicrosoftClientId, 'api-client-id');
      assert.equal(config.oauthBridgeMicrosoftClientSecret, 'api-client-secret');
      assert.doesNotThrow(() => requireRemoteConfigured(config));
    }
  );
});

test('a separate confidential web client requires its matching secret', () => {
  withEnvironment(
    {
      ...common,
      MCP_OAUTH_BRIDGE_MICROSOFT_CLIENT_TYPE: 'confidential_web',
      MS_OAUTH_CLIENT_ID: 'separate-web-client-id'
    },
    () => {
      const config = loadConfig();
      assert.equal(config.oauthBridgeMicrosoftClientId, 'separate-web-client-id');
      assert.equal(config.oauthBridgeMicrosoftClientSecret, undefined);
      assert.throws(() => requireRemoteConfigured(config), /MS_OAUTH_CLIENT_SECRET/);
    }
  );
});

test('invalid Microsoft client type fails closed', () => {
  withEnvironment(
    {
      ...common,
      MCP_OAUTH_BRIDGE_MICROSOFT_CLIENT_TYPE: 'desktop_web_mix'
    },
    () => {
      assert.throws(() => loadConfig(), ConfigError);
    }
  );
});

test('send mode defaults to confirmation and can be switched to automatic', () => {
  withEnvironment(common, () => {
    assert.equal(loadConfig().sendMode, 'confirm');
  });
  withEnvironment({ ...common, MCP_SEND_MODE: 'automatic' }, () => {
    assert.equal(loadConfig().sendMode, 'automatic');
  });
});

test('invalid send mode fails closed', () => {
  withEnvironment({ ...common, MCP_SEND_MODE: 'agent_decides' }, () => {
    assert.throws(() => loadConfig(), ConfigError);
  });
});

test('tool exposure mode and direct fast lane lists are parsed strictly', () => {
  withEnvironment(
    {
      ...common,
      MCP_TOOL_EXPOSURE_MODE: 'hybrid',
      MCP_DIRECT_TOOL_CATEGORIES: 'smart, mail',
      MCP_DIRECT_TOOLS: 'graph_get_me,calendar_list_events'
    },
    () => {
      const config = loadConfig();
      assert.equal(config.toolExposureMode, 'hybrid');
      assert.deepEqual(config.directToolCategories, ['smart', 'mail']);
      assert.deepEqual(config.directTools, ['graph_get_me', 'calendar_list_events']);
    }
  );
});

test('invalid tool exposure mode fails closed', () => {
  withEnvironment({ ...common, MCP_TOOL_EXPOSURE_MODE: 'auto' }, () => {
    assert.throws(() => loadConfig(), ConfigError);
  });
});
