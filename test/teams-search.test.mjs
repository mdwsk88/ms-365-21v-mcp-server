import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';

import { defaultDirectTools, loadConfig } from '../dist/config.js';
import { listChannelMessagesOverview, searchTeamsMessages } from '../dist/graph-scenarios.js';
import { runWithRequestContext } from '../dist/request-context.js';

const graph = 'https://microsoftgraph.chinacloudapi.cn/v1.0';

async function graphTestConfig(t) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'ms365-21v-teams-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
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
    ...loadConfig(),
    tokenCachePath,
    deviceCodeCachePath: `${tokenCachePath}.device-code`,
    graphBaseUrl: graph,
    graphResource: 'https://microsoftgraph.chinacloudapi.cn',
    auditLogEnabled: false,
    disabledGraphScopes: [],
    allowedGraphScopes: [],
    graphResilience: {
      maxRetries: 0,
      initialBackoffMs: 0,
      backoffMultiplier: 1,
      circuitBreakerThreshold: 100,
      circuitBreakerCooldownMs: 1000,
      timeoutMs: 1000
    }
  };
}

const json = body => new Response(JSON.stringify(body), { status: 200, headers: { 'content-type': 'application/json' } });

async function withMockFetch(handler, callback) {
  const original = globalThis.fetch;
  const requests = [];
  globalThis.fetch = async (url, init = {}) => {
    const request = { url: new URL(String(url)), init };
    requests.push(request);
    return handler(request);
  };
  try {
    const result = await runWithRequestContext(
      { requestId: 'teams-test', inboundClaims: { oid: 'me-id', preferred_username: 'me@example.cn' } },
      callback
    );
    return { result, requests };
  } finally {
    globalThis.fetch = original;
  }
}

test('Teams search returns compact hits that point to their chat or channel', async (t) => {
  const config = await graphTestConfig(t);
  const { result, requests } = await withMockFetch(
    () =>
      json({
        value: [
          {
            searchTerms: ['budget'],
            hitsContainers: [
              {
                total: 42,
                moreResultsAvailable: true,
                hits: [
                  {
                    hitId: 'h1',
                    summary: 'Q4 <c0>budget</c0> &amp; plan<ddd/>',
                    resource: {
                      '@odata.type': '#microsoft.graph.chatMessage',
                      id: '1700000000001',
                      createdDateTime: '2026-11-02T08:00:00Z',
                      chatId: '19:abc@unq.gbl.spaces',
                      from: { emailAddress: { name: 'Zhang Wei', address: 'zhang.wei@example.cn' } },
                      etag: 'e1'
                    }
                  },
                  {
                    hitId: 'h2',
                    summary: '<c0>Budget</c0> review',
                    resource: {
                      id: '1700000000002',
                      createdDateTime: '2026-11-01T08:00:00Z',
                      channelIdentity: { teamId: 'team-1', channelId: '19:general@thread.tacv2' },
                      subject: 'Finance',
                      from: { emailAddress: { name: 'Li Na' } }
                    }
                  }
                ]
              }
            ]
          }
        ]
      }),
    () => searchTeamsMessages(config, { query: 'budget', from: 20, size: 2 })
  );

  assert.equal(requests[0].url.pathname, '/v1.0/search/query');
  assert.deepEqual(JSON.parse(requests[0].init.body).requests[0], {
    entityTypes: ['chatMessage'],
    query: { queryString: 'budget' },
    from: 20,
    size: 2
  });
  assert.deepEqual(result, {
    query: 'budget',
    total: 42,
    count: 2,
    messages: [
      {
        id: '1700000000001',
        chatId: '19:abc@unq.gbl.spaces',
        teamId: undefined,
        channelId: undefined,
        from: 'Zhang Wei',
        at: '2026-11-02T08:00:00Z',
        subject: undefined,
        summary: 'Q4 budget & plan…',
        webLink: undefined
      },
      {
        id: '1700000000002',
        chatId: undefined,
        teamId: 'team-1',
        channelId: '19:general@thread.tacv2',
        from: 'Li Na',
        at: '2026-11-01T08:00:00Z',
        subject: 'Finance',
        summary: 'Budget review',
        webLink: undefined
      }
    ],
    nextFrom: 22
  });
});

test('channel posts and replies come back as compact messages, oldest first', async (t) => {
  const config = await graphTestConfig(t);
  const post = (id, at, extra = {}) => ({
    id,
    createdDateTime: at,
    messageType: 'message',
    from: { user: { id: id === 'p2' ? 'me-id' : 'u1', displayName: id === 'p2' ? 'Me' : 'Wang' } },
    body: { contentType: 'html', content: `<p>Post ${id}</p>` },
    webUrl: `https://teams.microsoftonline.cn/l/message/${id}`,
    ...extra
  });
  const nextLink = `${graph}/teams/team-1/channels/19%3Ageneral%40thread.tacv2/messages?%24skiptoken=abc`;
  const posts = await withMockFetch(
    () =>
      json({
        value: [
          post('p2', '2026-11-02T09:00:00Z', { subject: 'Release' }),
          post('p1', '2026-11-01T09:00:00Z'),
          post('p0', '2026-10-31T09:00:00Z', { deletedDateTime: '2026-11-01T00:00:00Z' })
        ],
        '@odata.nextLink': nextLink
      }),
    () => listChannelMessagesOverview(config, { teamId: 'team-1', channelId: '19:general@thread.tacv2' })
  );
  assert.equal(posts.requests[0].url.pathname, '/v1.0/teams/team-1/channels/19%3Ageneral%40thread.tacv2/messages');
  assert.equal(posts.requests[0].url.searchParams.get('$top'), '20');
  assert.deepEqual(
    posts.result.messages.map(message => [message.id, message.text, message.fromMe, message.subject]),
    [
      ['p1', 'Post p1', undefined, undefined],
      ['p2', 'Post p2', true, 'Release']
    ]
  );
  assert.equal(posts.result.messages[0].webLink, 'https://teams.microsoftonline.cn/l/message/p1');
  assert.equal(posts.result.nextCursor, nextLink);

  const replies = await withMockFetch(
    () => json({ value: [post('r1', '2026-11-02T10:00:00Z')] }),
    () => listChannelMessagesOverview(config, { teamId: 'team-1', channelId: 'c1', messageId: 'p2', top: 5 })
  );
  assert.equal(replies.requests[0].url.pathname, '/v1.0/teams/team-1/channels/c1/messages/p2/replies');
  assert.equal(replies.result.messageId, 'p2');

  const page = await withMockFetch(
    () => json({ value: [] }),
    () => listChannelMessagesOverview(config, { teamId: 'team-1', channelId: 'c1', cursor: nextLink })
  );
  assert.equal(page.requests[0].url.href, nextLink);
  await assert.rejects(
    () => listChannelMessagesOverview(config, { teamId: 'team-1', channelId: 'c1', cursor: `${graph}/me/messages?$skip=1` }),
    /different resource/
  );
});

test('Teams search stays out of the direct tools until it is verified in 21Vianet', () => {
  assert.equal(defaultDirectTools.includes('search_teams'), false);
});
