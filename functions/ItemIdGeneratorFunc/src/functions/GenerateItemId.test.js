'use strict';

const assert = require('node:assert/strict');
const { afterEach, describe, it } = require('node:test');
const { generateItemId, registerRule } = require('../rules/idRules');
const { notifyWebhook } = require('../webhook/notifyWebhook');
const { generateItemIdHandler } = require('./GenerateItemId');

const fixedNow = new Date(1700000000000);

function requestWith(body) {
  return {
    json: async () => body
  };
}

function contextStub() {
  return {
    errors: [],
    error(message) {
      this.errors.push(message);
    }
  };
}

describe('generateItemId rules', { concurrency: false }, () => {
  it('uses list name and timestamp by default', () => {
    const id = generateItemId(
      { listName: 'Requests', condition: 'default', itemId: 15 },
      fixedNow
    );
    assert.equal(id, 'Requests-1700000000000');
  });

  it('prefixes VIP when the condition is VIP, ignoring case and surrounding spaces', () => {
    const id = generateItemId(
      { listName: 'HR Requests', condition: '  vip ', itemId: 4 },
      fixedNow
    );
    assert.equal(id, 'VIP-HR-Requests-1700000000000');
  });

  it('does not treat a partial VIP condition as the VIP rule', () => {
    const id = generateItemId(
      { listName: 'Requests', condition: 'VIP-guest', itemId: 4 },
      fixedNow
    );
    assert.equal(id, 'Requests-1700000000000');
  });

  it('lets a registered rule read extra metadata and take priority', () => {
    const unregister = registerRule({
      id: 'department',
      matches: (ctx) => typeof ctx.metadata.department === 'string' && ctx.metadata.department.trim() !== '',
      format: (ctx) => `${ctx.metadata.department.trim()}-${ctx.listName}-${ctx.timestamp}`
    });

    try {
      const id = generateItemId(
        {
          listName: 'Requests',
          condition: 'VIP',
          itemId: 9,
          metadata: { department: 'Finance' }
        },
        fixedNow
      );
      assert.equal(id, 'Finance-Requests-1700000000000');
    } finally {
      unregister();
    }

    const after = generateItemId(
      { listName: 'Requests', condition: 'VIP', itemId: 9, metadata: { department: 'Finance' } },
      fixedNow
    );
    assert.equal(after, 'VIP-Requests-1700000000000');
  });
});

describe('notifyWebhook', { concurrency: false }, () => {
  it('skips when WEBHOOK_URL is unset', async () => {
    let called = false;
    const result = await notifyWebhook(
      { id: 'Requests-1', itemId: 1, condition: 'default', listName: 'Requests' },
      {
        env: {},
        fetchImpl: async () => {
          called = true;
          return { ok: true, status: 200 };
        }
      }
    );
    assert.equal(result.skipped, true);
    assert.equal(called, false);
  });

  it('posts the generated id payload', async () => {
    const calls = [];
    const result = await notifyWebhook(
      { id: 'VIP-Requests-1', itemId: 8, condition: 'VIP', listName: 'Requests' },
      {
        env: { WEBHOOK_URL: ' https://webhook.example/hook ' },
        fetchImpl: async (url, options) => {
          calls.push({ url, options });
          return { ok: true, status: 202 };
        }
      }
    );

    assert.equal(result.skipped, false);
    assert.equal(result.status, 202);
    assert.equal(calls[0].url, 'https://webhook.example/hook');
    assert.equal(calls[0].options.method, 'POST');
    assert.deepEqual(JSON.parse(calls[0].options.body), {
      id: 'VIP-Requests-1',
      itemId: 8,
      condition: 'VIP',
      listName: 'Requests'
    });
  });
});

describe('GenerateItemId handler', { concurrency: false }, () => {
  const originalFetch = globalThis.fetch;
  const originalWebhook = process.env.WEBHOOK_URL;

  afterEach(() => {
    globalThis.fetch = originalFetch;
    if (originalWebhook === undefined) {
      delete process.env.WEBHOOK_URL;
    } else {
      process.env.WEBHOOK_URL = originalWebhook;
    }
  });

  it('returns HTTP 200 and JSON body { id } for a VIP request', async () => {
    delete process.env.WEBHOOK_URL;
    const result = await generateItemIdHandler(
      requestWith({ listName: 'Requests', condition: 'VIP', itemId: 15, department: 'Finance' }),
      contextStub(),
      { generateItemId: (input) => generateItemId(input, fixedNow) }
    );

    assert.equal(result.status, 200);
    assert.deepEqual(result.jsonBody, { id: 'VIP-Requests-1700000000000' });
  });

  it('returns 400 when the body is not JSON', async () => {
    const result = await generateItemIdHandler(
      {
        json: async () => {
          throw new SyntaxError('Unexpected token');
        }
      },
      contextStub()
    );
    assert.equal(result.status, 400);
    assert.equal(result.jsonBody.error, 'Request body must be JSON.');
  });

  it('returns 400 when itemId is missing', async () => {
    const result = await generateItemIdHandler(
      requestWith({ listName: 'Requests', condition: 'default' }),
      contextStub()
    );
    assert.equal(result.status, 400);
    assert.match(result.jsonBody.error, /itemId/);
  });

  it('still returns the id when the webhook fails', async () => {
    process.env.WEBHOOK_URL = 'https://webhook.example/hook';
    globalThis.fetch = async () => ({ ok: false, status: 502 });
    const context = contextStub();

    const result = await generateItemIdHandler(
      requestWith({ listName: 'Requests', condition: 'default', itemId: '15' }),
      context,
      { generateItemId: (input) => generateItemId(input, fixedNow) }
    );

    assert.equal(result.status, 200);
    assert.deepEqual(result.jsonBody, { id: 'Requests-1700000000000' });
    assert.equal(context.errors.length, 1);
    assert.match(context.errors[0], /webhook/i);
  });
});
