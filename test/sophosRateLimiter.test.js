const test = require('node:test');
const assert = require('node:assert/strict');

const { createSophosRateLimiter, runSophosRequestWithRetry, shouldProcessActionableAlerts, getSophosSiemAlerts, isFreshSophosMetadataCache, shouldRunClosedAlertsCheck, isFreshAutotaskLocationCacheEntry, isFreshDeviceCacheEntry, updateSophosAlertQueryFailureState, getAutotaskDeviceCacheKey, deduplicateAlerts, getAutotaskTicketSearchKey, getAdaptiveSophosRetryMaxAttempts, shouldStopSophosTenantLoop } = require('../SophosAlerts_AutotaskIntegration/index.js');

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

test('Adaptive retry ceiling expands after a healthy recent success streak', () => {
  const state = { recentSuccesses: 4, consecutive429s: 0, recent429s: 0 };

  assert.equal(getAdaptiveSophosRetryMaxAttempts(state, { baseMaxAttempts: 3, maxSafeAttempts: 8 }), 5);
});

test('Adaptive retry stops tenant processing when throttling remains sustained', () => {
  const state = { recentSuccesses: 0, consecutive429s: 4, recent429s: 4 };

  assert.equal(shouldStopSophosTenantLoop(state), true);
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

test('Autotask location cache remains valid for seven days', () => {
  const now = Date.parse('2026-09-14T12:00:00.000Z');
  const baseEntry = {
    cachedAt: '2026-09-07T12:00:01.000Z',
    location: null
  };

  assert.equal(isFreshAutotaskLocationCacheEntry(baseEntry, now), true);
  assert.equal(isFreshAutotaskLocationCacheEntry({ ...baseEntry, cachedAt: '2026-09-07T12:00:00.000Z' }, now), false);
  assert.equal(isFreshAutotaskLocationCacheEntry({ cachedAt: baseEntry.cachedAt }, now), false);
  assert.equal(isFreshAutotaskLocationCacheEntry({ ...baseEntry, cachedAt: '2026-09-14T13:00:00.000Z' }, now), false);
});

test('Device caches expire after 24 hours and Autotask keys include device identity', () => {
  const now = Date.parse('2026-09-14T12:00:00.000Z');
  const entry = { cachedAt: '2026-09-13T12:00:01.000Z', deviceID: 123 };
  const deviceDetails = {
    hostname: 'workstation-1',
    macAddresses: ['BB', 'AA'],
    associatedPerson: { viaLogin: 'user@example.com' },
    ipv4Addresses: ['10.0.0.2']
  };

  assert.equal(isFreshDeviceCacheEntry(entry, now), true);
  assert.equal(isFreshDeviceCacheEntry({ ...entry, cachedAt: '2026-09-13T12:00:00.000Z' }, now), false);
  assert.equal(isFreshDeviceCacheEntry({ cachedAt: entry.cachedAt }, now), true);
  assert.notEqual(getAutotaskDeviceCacheKey(42, deviceDetails), getAutotaskDeviceCacheKey(43, deviceDetails));
  assert.equal(
    getAutotaskDeviceCacheKey(42, deviceDetails),
    getAutotaskDeviceCacheKey(42, { ...deviceDetails, macAddresses: ['AA', 'BB'] })
  );
});

test('Sophos alert query failures are counted within a rolling 24-hour window', () => {
  const now = Date.parse('2026-10-05T12:00:00.000Z');
  let state = null;

  for (let failure = 1; failure <= 4; failure++) {
    state = updateSophosAlertQueryFailureState(state, now + (failure * 1000));
    assert.equal(state.failureCount, failure);
  }

  state = updateSophosAlertQueryFailureState(state, now + 5000);
  assert.equal(state.failureCount, 5);
  assert.equal(state.lastFailureAt, new Date(now + 5000).toISOString());

  const expiredState = updateSophosAlertQueryFailureState(state, now + (24 * 60 * 60 * 1000) + 5000);
  assert.equal(expiredState.failureCount, 1);
});

test('Alert deduplication keeps distinct alerts and removes repeated IDs', () => {
  const alerts = [
    { id: 'alert-1', severity: 'high' },
    { id: 'alert-1', severity: 'high' },
    { id: 'alert-2', severity: 'medium' },
    { severity: 'high' },
    { severity: 'high' }
  ];

  assert.deepEqual(deduplicateAlerts(alerts).map(alert => alert.id), ['alert-1', 'alert-2', undefined, undefined]);
});

test('Autotask ticket search keys distinguish search scope', () => {
  const baseKey = getAutotaskTicketSearchKey(42, 'Sophos Alert: ', 'device-1', 'event-1');
  assert.equal(baseKey, getAutotaskTicketSearchKey(42, 'Sophos Alert: ', 'device-1', 'event-1'));
  assert.notEqual(baseKey, getAutotaskTicketSearchKey(42, 'Sophos Alert: ', 'device-2', 'event-1'));
  assert.notEqual(baseKey, getAutotaskTicketSearchKey(43, 'Sophos Alert: ', 'device-1', 'event-1'));
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
