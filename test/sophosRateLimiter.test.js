const test = require('node:test');
const assert = require('node:assert/strict');

const { createSophosRateLimiter, runSophosRequestWithRetry, shouldProcessActionableAlerts, getSophosSiemAlerts, isFreshSophosMetadataCache, shouldRunClosedAlertsCheck } = require('../SophosAlerts_AutotaskIntegration/index.js');

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

test('Empty or non-actionable alert sets should not continue processing', () => {
  assert.equal(shouldProcessActionableAlerts([], []), false, 'No alerts should be skipped');
  assert.equal(shouldProcessActionableAlerts([], [{ severity: 'low', type: 'Event::Endpoint::ServiceRestored' }]), true, 'Matching up/down events should continue');
  assert.equal(shouldProcessActionableAlerts([{ severity: 'medium' }], []), true, 'Medium alerts should continue');
  assert.equal(shouldProcessActionableAlerts([], [{ severity: 'low', type: 'not-a-real-event' }]), false, 'Unknown low-severity events should be skipped');
});

test('Sophos metadata cache is fresh for less than 24 hours only', () => {
  const now = Date.parse('2026-09-14T12:00:00.000Z');
  const baseCache = {
    cachedAt: '2026-09-14T00:00:00.000Z',
    partnerID: 'partner-1',
    tenants: { items: [{ id: 'tenant-1', status: 'active' }] }
  };

  assert.equal(isFreshSophosMetadataCache(baseCache, now), true);
  assert.equal(isFreshSophosMetadataCache({ ...baseCache, cachedAt: '2026-09-13T12:00:00.000Z' }, now), false);
  assert.equal(isFreshSophosMetadataCache({ ...baseCache, tenants: null }, now), false);
});

test('Closed-alert sweep runs immediately and then every two hours', () => {
  const now = Date.parse('2026-09-14T12:00:00.000Z');

  assert.equal(shouldRunClosedAlertsCheck(null, now), true);
  assert.equal(shouldRunClosedAlertsCheck(new Date('2026-09-14T10:00:00.000Z'), now), true);
  assert.equal(shouldRunClosedAlertsCheck(new Date('2026-09-14T10:00:01.000Z'), now), false);
  assert.equal(shouldRunClosedAlertsCheck(new Date('2026-09-14T12:01:00.000Z'), now), false);
});

test('Sophos alert query failures should be surfaced instead of silently writing the checkpoint', async () => {
  const originalFetch = global.fetch;
  const originalAxiosGet = require('axios').get;

  global.fetch = async () => {
    throw new Error('network failure');
  };
  require('axios').get = async () => {
    throw new Error('network failure');
  };

  try {
    await assert.rejects(
      () => getSophosSiemAlerts(
        { log: () => {}, warn: () => {}, error: () => {} },
        'token',
        { items: [{ id: 'tenant-1', status: 'active', dataRegion: 'example' }] },
        123
      ),
      /Sophos alerts query failed/
    );
  } finally {
    global.fetch = originalFetch;
    require('axios').get = originalAxiosGet;
  }
});
