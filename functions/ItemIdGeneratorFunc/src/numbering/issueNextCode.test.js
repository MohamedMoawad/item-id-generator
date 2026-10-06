'use strict';

const assert = require('node:assert/strict');
const { describe, it } = require('node:test');
const { issueNextCode } = require('./issueNextCode');

class InMemoryConfigStore {
  constructor(fields) {
    this.version = 1;
    this.fields = {
      currentCount: 0,
      formula: 'REQ-{seq:4}',
      resetPeriod: 'None',
      isActive: true,
      lastResetDate: '',
      padLength: 4,
      ...fields
    };
  }

  async read() {
    await Promise.resolve();
    return { ...this.fields, etag: String(this.version) };
  }

  async compareAndSwap(etag, patch) {
    await Promise.resolve();
    if (String(this.version) !== String(etag)) {
      return { ok: false };
    }
    this.version += 1;
    this.fields = {
      ...this.fields,
      currentCount: patch.currentCount,
      lastResetDate: patch.lastResetDate
    };
    return { ok: true };
  }
}

const noWait = async () => {};

describe('issueNextCode', { concurrency: false }, () => {
  it('gives 100 parallel callers 100 distinct sequences', async () => {
    const store = new InMemoryConfigStore({ formula: 'REQ-{seq:4}', resetPeriod: 'None' });
    const now = new Date('2026-10-06T12:00:00Z');
    const issued = await Promise.all(
      Array.from({ length: 100 }, () => issueNextCode(store, now, { sleep: noWait, maxAttempts: 300 }))
    );
    const sequences = new Set(issued.map((result) => result.sequence));
    const codes = new Set(issued.map((result) => result.code));
    assert.equal(sequences.size, 100);
    assert.equal(codes.size, 100);
    assert.equal(Math.min(...sequences), 1);
    assert.equal(Math.max(...sequences), 100);
    assert.ok(codes.has('REQ-0001'));
    assert.ok(codes.has('REQ-0100'));
    const finalRow = await store.read();
    assert.equal(finalRow.currentCount, 100);
  });

  it('resets the count when the UTC month changes and then continues', async () => {
    const store = new InMemoryConfigStore({
      formula: 'REQ-{yyyy}{MM}-{seq:3}',
      resetPeriod: 'Month',
      currentCount: 40,
      lastResetDate: '2026-09-15T00:00:00.000Z'
    });
    const first = await issueNextCode(store, new Date('2026-10-06T00:00:00Z'), { sleep: noWait });
    const second = await issueNextCode(store, new Date('2026-10-06T01:00:00Z'), { sleep: noWait });
    assert.equal(first.code, 'REQ-202610-001');
    assert.equal(first.sequence, 1);
    assert.equal(first.lastResetDate, '2026-10-06T00:00:00.000Z');
    assert.equal(second.sequence, 2);
    assert.equal(second.code, 'REQ-202610-002');
  });

  it('resets on a UTC day or year boundary and does not reset when the period is None', async () => {
    const daily = new InMemoryConfigStore({
      formula: 'D-{seq}',
      resetPeriod: 'Day',
      currentCount: 9,
      lastResetDate: '2026-10-05T12:00:00.000Z',
      padLength: 0
    });
    const nextDay = await issueNextCode(daily, new Date('2026-10-06T00:00:00Z'), { sleep: noWait });
    assert.equal(nextDay.sequence, 1);
    assert.equal(nextDay.code, 'D-1');

    const yearly = new InMemoryConfigStore({
      formula: 'Y-{seq}',
      resetPeriod: 'Year',
      currentCount: 8,
      lastResetDate: '2025-06-01T00:00:00.000Z',
      padLength: 0
    });
    const nextYear = await issueNextCode(yearly, new Date('2026-01-01T00:00:00Z'), { sleep: noWait });
    assert.equal(nextYear.sequence, 1);

    const never = new InMemoryConfigStore({
      formula: 'N-{seq}',
      resetPeriod: 'None',
      currentCount: 5,
      lastResetDate: '2026-09-01T00:00:00.000Z',
      padLength: 0
    });
    const kept = await issueNextCode(never, new Date('2027-01-01T00:00:00Z'), { sleep: noWait });
    assert.equal(kept.sequence, 6);
  });

  it('does not wipe CurrentCount when LastResetDate is empty', async () => {
    const store = new InMemoryConfigStore({
      formula: 'REQ-{seq}',
      resetPeriod: 'Month',
      currentCount: 5,
      lastResetDate: '',
      padLength: 0
    });
    const issued = await issueNextCode(store, new Date('2026-10-06T00:00:00Z'), { sleep: noWait });
    assert.equal(issued.sequence, 6);
    assert.equal(issued.lastResetDate, '2026-10-06T00:00:00.000Z');
  });

  it('refuses an inactive row', async () => {
    const store = new InMemoryConfigStore({ isActive: false });
    await assert.rejects(
      () => issueNextCode(store, new Date('2026-10-06T00:00:00Z'), { sleep: noWait }),
      /inactive/
    );
  });
});
