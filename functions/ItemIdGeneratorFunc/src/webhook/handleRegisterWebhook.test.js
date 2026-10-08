'use strict';

const assert = require('node:assert/strict');
const { describe, it } = require('node:test');
const { handleRegisterWebhook } = require('./handleRegisterWebhook');

const LIST_ID = 'aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee';
const ENV = {
  NUMBERING_CONFIG_SITE_URL: 'https://contoso.sharepoint.com/sites/config',
  NUMBERING_CONFIG_LIST_TITLE: 'AutoGenFeatureConfiguration',
  SHAREPOINT_TENANT_ID: '11111111-2222-4333-8444-555555555555',
  SHAREPOINT_CLIENT_ID: '66666666-7777-4888-8999-aaaaaaaaaaaa',
  SHAREPOINT_CLIENT_SECRET: 'not-a-real-secret',
  SPO_WEBHOOK_NOTIFICATION_URL: 'https://contoso.azurewebsites.net/api/spoWebhook?code=not-a-real-key',
  SHAREPOINT_WEBHOOK_CLIENT_STATE: 'not-a-real-client-state'
};

function jsonResponse(status, body) {
  return {
    ok: status >= 200 && status < 300,
    status,
    async text() {
      return typeof body === 'string' ? body : JSON.stringify(body);
    }
  };
}

function requestWith(value) {
  return {
    async json() {
      if (value instanceof Error) {
        throw value;
      }
      return value;
    }
  };
}

describe('handleRegisterWebhook', () => {
  it('rejects a body that is not a config item id', async () => {
    const response = await handleRegisterWebhook(requestWith({ configItemId: 0 }), { log() {} }, { env: ENV });
    assert.equal(response.status, 400);
  });

  it('returns 404 when the config row is missing', async () => {
    const fetchImpl = async (url) => {
      if (url.indexOf('oauth2') >= 0) {
        return jsonResponse(200, { access_token: 'token', expires_in: 3600 });
      }
      return jsonResponse(404, '');
    };
    const response = await handleRegisterWebhook(
      requestWith({ configItemId: 7, configSiteUrl: 'https://contoso.sharepoint.com/sites/config' }),
      { log() {} },
      { env: ENV, fetchImpl }
    );
    assert.equal(response.status, 404);
  });

  it('registers the webhook for an active row', async () => {
    const fetchImpl = async (url, options = {}) => {
      if (url.indexOf('oauth2') >= 0) {
        return jsonResponse(200, { access_token: 'token', expires_in: 3600 });
      }
      if (url.endsWith('/subscriptions')) {
        return jsonResponse(201, { Id: 'sub-from-handler' });
      }
      if (options.headers && options.headers['X-HTTP-Method'] === 'MERGE') {
        return jsonResponse(204, '');
      }
      return jsonResponse(200, {
        Id: 7,
        TargetListUrl: 'https://contoso.sharepoint.com/sites/ops/Lists/Requests',
        TargetListGuid: LIST_ID,
        IsActive: true,
        WebhookSubscriptionId: '',
        'odata.etag': '"1"'
      });
    };
    const response = await handleRegisterWebhook(
      requestWith({ configItemId: 7, configSiteUrl: 'https://contoso.sharepoint.com/sites/other' }),
      { log() {} },
      { env: ENV, fetchImpl }
    );
    assert.equal(response.status, 200);
    assert.deepEqual(response.jsonBody, { subscriptionId: 'sub-from-handler', alreadyRegistered: false });
  });

  it('returns 502 when the row is inactive', async () => {
    const fetchImpl = async (url) => {
      if (url.indexOf('oauth2') >= 0) {
        return jsonResponse(200, { access_token: 'token', expires_in: 3600 });
      }
      return jsonResponse(200, {
        Id: 7,
        IsActive: false,
        TargetListUrl: 'https://contoso.sharepoint.com/sites/ops/Lists/Requests',
        TargetListGuid: LIST_ID,
        'odata.etag': '"1"'
      });
    };
    const response = await handleRegisterWebhook(
      requestWith({ configItemId: 7, configSiteUrl: 'https://contoso.sharepoint.com/sites/config' }),
      { log() {} },
      { env: ENV, fetchImpl }
    );
    assert.equal(response.status, 502);
    assert.match(response.jsonBody.error, /Activate/);
  });
});
