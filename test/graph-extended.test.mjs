import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';

import { loadConfig } from '../dist/config.js';
import {
  addChatMember,
  copyDriveItem,
  createTeamsChat,
  downloadDriveFile,
  getSharePointSiteByPath,
  searchMicrosoft365
} from '../dist/graph-extended.js';

async function graphTestConfig() {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'ms365-21v-graph-'));
  const tokenCachePath = path.join(directory, 'token.json');
  await fs.writeFile(
    tokenCachePath,
    JSON.stringify({
      accessToken: 'test-graph-token',
      tokenType: 'Bearer',
      expiresAt: Math.floor(Date.now() / 1000) + 3600,
      acquiredAt: Math.floor(Date.now() / 1000)
    }),
    { mode: 0o600 }
  );

  return {
    directory,
    config: {
      ...loadConfig(),
      tokenCachePath,
      deviceCodeCachePath: `${tokenCachePath}.device-code`,
      graphBaseUrl: 'https://microsoftgraph.chinacloudapi.cn/v1.0',
      graphResource: 'https://microsoftgraph.chinacloudapi.cn',
      auditLogEnabled: false,
      graphResilience: {
        maxRetries: 0,
        initialBackoffMs: 0,
        backoffMultiplier: 1,
        circuitBreakerThreshold: 100,
        circuitBreakerCooldownMs: 1000,
        timeoutMs: 1000
      }
    }
  };
}

async function withMockFetch(handler, callback) {
  const original = globalThis.fetch;
  globalThis.fetch = handler;
  try {
    return await callback();
  } finally {
    globalThis.fetch = original;
  }
}

test('Teams chat writes use China Graph bindings and valid chat roles', async (t) => {
  const { config, directory } = await graphTestConfig();
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const requests = [];

  await withMockFetch(
    async (url, init) => {
      requests.push({ url: String(url), init });
      return new Response('{"id":"chat-1"}', { status: 201, headers: { 'content-type': 'application/json' } });
    },
    async () => {
      await createTeamsChat(config, 'oneOnOne', [
        { userId: 'employee-1' },
        { userId: 'guest-1', role: 'guest' }
      ]);
      await addChatMember(config, 'chat/with special:id', { userId: 'employee-2' });
    }
  );

  assert.equal(requests.length, 2);
  assert.equal(requests[0].url, 'https://microsoftgraph.chinacloudapi.cn/v1.0/chats');
  assert.equal(requests[0].init.method, 'POST');
  assert.equal(requests[0].init.headers.Authorization, 'Bearer test-graph-token');
  const createBody = JSON.parse(requests[0].init.body);
  assert.deepEqual(createBody.members.map((member) => member.roles), [['owner'], ['guest']]);
  assert.equal(
    createBody.members[0]['user@odata.bind'],
    "https://microsoftgraph.chinacloudapi.cn/v1.0/users('employee-1')"
  );

  assert.equal(
    requests[1].url,
    'https://microsoftgraph.chinacloudapi.cn/v1.0/chats/chat%2Fwith%20special%3Aid/members'
  );
  assert.deepEqual(JSON.parse(requests[1].init.body).roles, ['owner']);
});

test('SharePoint path lookup and Microsoft Search encode the expected v1.0 requests', async (t) => {
  const { config, directory } = await graphTestConfig();
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const requests = [];

  await withMockFetch(
    async (url, init) => {
      requests.push({ url: String(url), init });
      return new Response('{"value":[]}', { status: 200, headers: { 'content-type': 'application/json' } });
    },
    async () => {
      await getSharePointSiteByPath(config, 'tenant.sharepoint.cn', '/sites/R&D Portal');
      await searchMicrosoft365(config, 'event', 'quarterly review', ['Calendars.Read'], 10, 20);
    }
  );

  assert.equal(
    requests[0].url,
    'https://microsoftgraph.chinacloudapi.cn/v1.0/sites/tenant.sharepoint.cn:/sites/R%26D%20Portal?%24select=id%2Cname%2CdisplayName%2Cdescription%2CwebUrl%2CcreatedDateTime%2ClastModifiedDateTime%2CsiteCollection'
  );
  assert.equal(requests[1].url, 'https://microsoftgraph.chinacloudapi.cn/v1.0/search/query');
  assert.deepEqual(JSON.parse(requests[1].init.body), {
    requests: [
      {
        entityTypes: ['event'],
        query: { queryString: 'quarterly review' },
        from: 10,
        size: 20
      }
    ]
  });
});

test('binary downloads are returned as bounded base64 MCP payloads', async (t) => {
  const { config, directory } = await graphTestConfig();
  t.after(() => fs.rm(directory, { recursive: true, force: true }));

  const result = await withMockFetch(
    async (url, init) => {
      assert.equal(String(url), 'https://microsoftgraph.chinacloudapi.cn/v1.0/me/drive/items/file-1/content');
      assert.equal(init.headers.Accept, '*/*');
      return new Response(Uint8Array.from([0, 1, 2, 253, 254, 255]), {
        status: 200,
        headers: { 'content-type': 'application/octet-stream', 'content-length': '6' }
      });
    },
    () => downloadDriveFile(config, 'file-1')
  );

  assert.deepEqual(result, {
    contentType: 'application/octet-stream',
    size: 6,
    contentBase64: 'AAEC/f7/'
  });
});

test('asynchronous copy responses preserve the Graph operation monitor URL', async (t) => {
  const { config, directory } = await graphTestConfig();
  t.after(() => fs.rm(directory, { recursive: true, force: true }));

  const result = await withMockFetch(
    async () =>
      new Response(null, {
        status: 202,
        headers: { location: 'https://microsoftgraph.chinacloudapi.cn/v1.0/monitor/operation-1' }
      }),
    () => copyDriveItem(config, 'source', 'target', 'copy.docx')
  );

  assert.deepEqual(result, {
    ok: true,
    status: 202,
    location: 'https://microsoftgraph.chinacloudapi.cn/v1.0/monitor/operation-1'
  });
});
