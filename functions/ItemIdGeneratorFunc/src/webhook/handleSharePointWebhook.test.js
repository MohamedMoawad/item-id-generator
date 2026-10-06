'use strict';

const assert = require('node:assert/strict');
const { describe, it } = require('node:test');
const { handleSharePointWebhook } = require('./handleSharePointWebhook');

const LIST_ID = 'aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee';

function requestWith(body, query) {
  return {
    url: 'https://func.example/api/SharePointWebhook',
    query: {
      get(name) {
        return query && Object.prototype.hasOwnProperty.call(query, name) ? query[name] : null;
      }
    },
    json: async () => body
  };
}

describe('handleSharePointWebhook', () => {
  it('echoes the SharePoint validation token as plain text', async () => {
    const response = await handleSharePointWebhook(
      requestWith(undefined, { validationtoken: 'abc 123' }),
      { log() {} },
      { env: {} }
    );
    assert.equal(response.status, 200);
    assert.equal(response.body, 'abc 123');
    assert.match(response.headers['Content-Type'], /text\/plain/);
  });

  it('rejects notifications when the client state secret is not configured', async () => {
    const response = await handleSharePointWebhook(
      requestWith({ value: [] }),
      { log() {} },
      { env: { SHAREPOINT_WEBHOOK_CLIENT_STATE: '<webhook-client-state>' } }
    );
    assert.equal(response.status, 500);
  });

  it('rejects a notification whose clientState does not match', async () => {
    let enqueued = false;
    const response = await handleSharePointWebhook(
      requestWith({
        value: [{
          clientState: 'wrong',
          resource: LIST_ID,
          siteUrl: 'https://contoso.sharepoint.com/sites/ops'
        }]
      }),
      { log() {} },
      {
        env: { SHAREPOINT_WEBHOOK_CLIENT_STATE: 'expected-state' },
        enqueue: async () => {
          enqueued = true;
        }
      }
    );
    assert.equal(response.status, 403);
    assert.equal(enqueued, false);
  });

  it('enqueues one message per notification and returns 200', async () => {
    const queued = [];
    const response = await handleSharePointWebhook(
      requestWith({
        value: [{
          clientState: 'expected-state',
          resource: `{${LIST_ID.toUpperCase()}}`,
          siteUrl: 'https://contoso.sharepoint.com/sites/ops',
          subscriptionId: 'sub-1'
        }]
      }),
      { log() {} },
      {
        env: { SHAREPOINT_WEBHOOK_CLIENT_STATE: 'expected-state' },
        enqueue: async (messages) => {
          queued.push(...messages);
        }
      }
    );
    assert.equal(response.status, 200);
    assert.equal(response.jsonBody.accepted, 1);
    assert.deepEqual(queued, [{
      siteUrl: 'https://contoso.sharepoint.com/sites/ops',
      listId: LIST_ID,
      subscriptionId: 'sub-1'
    }]);
  });
});
