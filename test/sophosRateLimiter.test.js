const test = require('node:test');
const assert = require('node:assert/strict');

const { createSophosRateLimiter, runSophosRequestWithRetry } = require('../SophosAlerts_AutotaskIntegration/index.js');

test('Sophos rate limiter keeps requests at or below 10/sec', async () => {
  const limiter = createSophosRateLimiter({ log: () => {} }, 10);

  const start = Date.now();
  await limiter();
  await limiter();
  const elapsed = Date.now() - start;

  assert.ok(elapsed >= 95, `Expected at least 95ms between requests, got ${elapsed}ms`);
});

test('Sophos 429 retry loop is capped to a finite number of attempts', async () => {
  let attempts = 0;

  const result = await runSophosRequestWithRetry(
    { log: () => {} },
    '429 test',
    async () => {
      attempts += 1;
      const error = new Error('TooManyRequests');
      error.response = { status: 429 };
      throw error;
    },
    { maxAttempts: 3, backoffMs: 0 }
  );

  assert.equal(attempts, 3, `Expected 3 total attempts, got ${attempts}`);
  assert.equal(result, null, 'Expected null result after exhausting retries');
});
