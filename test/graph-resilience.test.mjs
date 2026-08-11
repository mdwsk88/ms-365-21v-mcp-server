import assert from 'node:assert/strict';
import test from 'node:test';

import { GraphCircuitOpenError, GraphRequestTimeoutError, GraphResilience } from '../dist/graph-resilience.js';

function resilienceConfig(overrides = {}) {
  return {
    maxRetries: 3,
    initialBackoffMs: 100,
    backoffMultiplier: 2,
    circuitBreakerThreshold: 2,
    circuitBreakerCooldownMs: 1_000,
    timeoutMs: 100,
    ...overrides
  };
}

test('429 honors Retry-After and retries the request', async () => {
  const delays = [];
  let discarded = false;
  const resilience = new GraphResilience(resilienceConfig(), {
    sleep: async (value) => delays.push(value)
  });
  let calls = 0;
  const response = await resilience.execute(async () => {
    calls += 1;
    return calls === 1
      ? new Response(
          new ReadableStream({
            start(controller) {
              controller.enqueue(new TextEncoder().encode('limited'));
            },
            cancel() {
              discarded = true;
            }
          }),
          { status: 429, headers: { 'Retry-After': '2' } }
        )
      : new Response('{}', { status: 200 });
  });
  assert.equal(response.status, 200);
  assert.equal(calls, 2);
  assert.deepEqual(delays, [2_000]);
  assert.equal(discarded, true);
});

test('503 and 504 use exponential backoff while 500 is never retried', async () => {
  const delays = [];
  const resilience = new GraphResilience(resilienceConfig(), {
    sleep: async (value) => delays.push(value)
  });
  const statuses = [503, 504, 200];
  const response = await resilience.execute(async () => new Response('{}', { status: statuses.shift() }));
  assert.equal(response.status, 200);
  assert.deepEqual(delays, [100, 200]);

  let calls = 0;
  const noRetry = new GraphResilience(resilienceConfig());
  const serverError = await noRetry.execute(async () => {
    calls += 1;
    return new Response('{}', { status: 500 });
  });
  assert.equal(serverError.status, 500);
  assert.equal(calls, 1);
});

test('unsafe operations can disable ambiguous network and 503 retries', async () => {
  let calls = 0;
  const resilience = new GraphResilience(resilienceConfig());
  const response = await resilience.execute(
    async () => {
      calls += 1;
      return new Response('{}', { status: 503 });
    },
    { retryTransientFailures: false }
  );
  assert.equal(response.status, 503);
  assert.equal(calls, 1);

  await assert.rejects(
    () =>
      resilience.execute(
        async () => {
          calls += 1;
          throw new Error('connection reset after write');
        },
        { retryTransientFailures: false }
      ),
    /connection reset after write/
  );
  assert.equal(calls, 2);
});

test('circuit breaker rejects requests until its cooldown elapses', async () => {
  let now = 10_000;
  const resilience = new GraphResilience(resilienceConfig({ maxRetries: 0 }), { now: () => now });
  await resilience.execute(async () => new Response('{}', { status: 500 }));
  await resilience.execute(async () => new Response('{}', { status: 500 }));
  await assert.rejects(
    () => resilience.execute(async () => new Response('{}', { status: 200 })),
    GraphCircuitOpenError
  );
  now += 1_001;
  const recovered = await resilience.execute(async () => new Response('{}', { status: 200 }));
  assert.equal(recovered.status, 200);
  assert.deepEqual(resilience.snapshot(), {
    consecutiveFailures: 0,
    circuitOpen: false,
    retryAfterMs: 0
  });
});

test('request timeout aborts a stalled Graph call', async () => {
  const resilience = new GraphResilience(resilienceConfig({ maxRetries: 0, timeoutMs: 5 }));
  await assert.rejects(() => resilience.execute(async () => new Promise(() => {})), GraphRequestTimeoutError);
});
