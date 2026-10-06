'use strict';

const assert = require('node:assert/strict');
const { describe, it } = require('node:test');
const { issueNextCode } = require('../numbering/issueNextCode');
const { processNumberingMessage } = require('./processNumberingMessage');

const LIST_ID = 'aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee';

class InMemoryConfigStore {
  constructor(fields) {
    this.version = 1;
    this.fields = {
      currentCount: 0,
      periodKey: '',
      formula: 'REQ-{seq:4}',
      resetPeriod: 'None',
      enabled: true,
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
    this.fields = { ...this.fields, ...patch };
    return { ok: true };
  }
}

class InMemoryLeaseStore {
  constructor() {
    this.held = new Set();
  }

  async acquire(key) {
    if (this.held.has(key)) {
      return { acquired: false };
    }
    this.held.add(key);
    return {
      acquired: true,
      release: async () => {
        this.held.delete(key);
      }
    };
  }
}

function createGateway(itemCount) {
  const store = new InMemoryConfigStore({});
  const items = Array.from({ length: itemCount }, (_, index) => ({ id: index + 1, number: '' }));
  return {
    items,
    async getEnabledConfig() {
      return {
        id: 7,
        targetSiteUrl: 'https://contoso.sharepoint.com/sites/ops',
        targetListId: LIST_ID,
        numberFieldInternalName: 'RequestNumber',
        enabled: true
      };
    },
    async listUnnumbered() {
      return items.filter((item) => !item.number).map((item) => ({ id: item.id }));
    },
    async reserveNextCode(_config, now) {
      return issueNextCode(store, now, { sleep: async () => {}, maxAttempts: 20 });
    },
    async writeNumberIfBlank(_config, itemId, code) {
      const item = items.find((entry) => entry.id === itemId);
      if (!item || item.number) {
        return { skipped: true };
      }
      item.number = code;
      return { skipped: false };
    }
  };
}

describe('processNumberingMessage', { concurrency: false }, () => {
  it('lets one leased worker number every blank item and makes the others retry', async () => {
    const gateway = createGateway(100);
    const leaseStore = new InMemoryLeaseStore();
    const message = {
      siteUrl: 'https://contoso.sharepoint.com/sites/ops',
      listId: `{${LIST_ID.toUpperCase()}}`
    };
    const results = await Promise.allSettled(
      Array.from({ length: 20 }, () => processNumberingMessage(message, { log() {} }, {
        leaseStore,
        gateway,
        now: () => new Date('2026-10-06T12:00:00Z')
      }))
    );

    const fulfilled = results.filter((result) => result.status === 'fulfilled');
    const rejected = results.filter((result) => result.status === 'rejected');
    assert.equal(fulfilled.length, 1);
    assert.equal(rejected.length, 19);
    assert.match(rejected[0].reason.message, /lease/i);
    const numbers = gateway.items.map((item) => item.number);
    assert.equal(new Set(numbers).size, 100);
    assert.equal(numbers[0], 'REQ-0001');
    assert.equal(numbers[99], 'REQ-0100');
  });

  it('acks a list that has no enabled config', async () => {
    const result = await processNumberingMessage(
      { siteUrl: 'https://contoso.sharepoint.com/sites/ops', listId: LIST_ID },
      { log() {} },
      {
        leaseStore: new InMemoryLeaseStore(),
        gateway: {
          async getEnabledConfig() {
            return undefined;
          }
        }
      }
    );
    assert.deepEqual(result.issued, []);
  });
});
