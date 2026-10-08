'use strict';

const assert = require('node:assert/strict');
const { describe, it } = require('node:test');
const { issueNextCode } = require('../numbering/issueNextCode');
const {
  createConfigStore,
  createSharePointDeps,
  getEnabledConfig,
  registerListWebhook
} = require('./sharePointClient');
const {
  buildRecentItemsUrl,
  deriveWebFromListUrl,
  isBlankNumber,
  listIdFromResource,
  normalizeGuid
} = require('./sharePointUrls');

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

function activeRow(overrides) {
  return {
    Id: 7,
    Title: 'Requests',
    TargetListUrl: 'https://contoso.sharepoint.com/sites/ops/Lists/Requests',
    TargetListGuid: LIST_ID,
    NumberColumnInternalName: 'RequestNumber',
    CurrentCount: 0,
    Formula: 'REQ-{seq:4}',
    ResetPeriod: 'None',
    LastResetDate: '',
    IsActive: true,
    PadLength: 4,
    WebhookSubscriptionId: '',
    'odata.etag': '"1"',
    ...overrides
  };
}

describe('SharePoint config updates', () => {
  it('retries the counter when If-Match loses and then returns the next code', async () => {
    const calls = [];
    let reads = 0;
    const fetchImpl = async (url, options = {}) => {
      calls.push({ url, options });
      if (url.indexOf('oauth2/v2.0/token') >= 0) {
        return jsonResponse(200, { access_token: 'token', expires_in: 3600 });
      }
      if (options.headers && options.headers['X-HTTP-Method'] === 'MERGE') {
        if (options.headers['IF-MATCH'] === '"1"') {
          return jsonResponse(412, '');
        }
        assert.equal(options.headers['IF-MATCH'], '"2"');
        assert.deepEqual(JSON.parse(options.body), {
          CurrentCount: 2,
          LastResetDate: '2026-10-06T12:00:00.000Z'
        });
        return jsonResponse(204, '');
      }
      reads += 1;
      return jsonResponse(200, activeRow({
        CurrentCount: reads === 1 ? 0 : 1,
        'odata.etag': reads === 1 ? '"1"' : '"2"'
      }));
    };

    const deps = createSharePointDeps(ENV, fetchImpl);
    const issued = await issueNextCode(
      createConfigStore(deps, { id: 7 }),
      new Date('2026-10-06T12:00:00.000Z'),
      { sleep: async () => {}, maxAttempts: 5 }
    );
    assert.equal(issued.code, 'REQ-0002');
    assert.equal(issued.sequence, 2);
    assert.ok(calls.some((call) => call.options.headers && call.options.headers['IF-MATCH'] === '"1"'));
  });

  it('matches the target list GUID and ignores an inactive row', async () => {
    const fetchImpl = async (url) => {
      if (url.indexOf('oauth2') >= 0) {
        return jsonResponse(200, { access_token: 'token', expires_in: 3600 });
      }
      return jsonResponse(200, {
        value: [
          activeRow({
            Id: 1,
            TargetListGuid: `{${LIST_ID.toUpperCase()}}`,
            IsActive: false,
            CurrentCount: 3
          }),
          activeRow({
            Id: 2,
            CurrentCount: 4,
            Formula: 'REQ-{seq}',
            ResetPeriod: 'Month',
            IsActive: 'Yes',
            'odata.etag': '"4"'
          })
        ]
      });
    };
    const config = await getEnabledConfig(createSharePointDeps(ENV, fetchImpl), LIST_ID);
    assert.equal(config.id, 2);
    assert.equal(config.currentCount, 4);
    assert.equal(config.targetListGuid, LIST_ID);
    assert.equal(config.targetSiteUrl, 'https://contoso.sharepoint.com/sites/ops');
    assert.equal(config.isActive, true);
  });
});

describe('registerListWebhook', () => {
  it('returns a stored subscription without calling SharePoint', async () => {
    const deps = createSharePointDeps(ENV, async () => {
      throw new Error('fetch should not run');
    });
    const result = await registerListWebhook(deps, {
      id: 3,
      isActive: true,
      webhookSubscriptionId: 'sub-existing',
      targetListUrl: 'https://contoso.sharepoint.com/sites/ops/Lists/Requests',
      targetListGuid: LIST_ID,
      etag: '"1"'
    });
    assert.deepEqual(result, { subscriptionId: 'sub-existing', alreadyRegistered: true });
  });

  it('rejects an inactive row and a placeholder notification URL', async () => {
    const deps = createSharePointDeps(ENV, async () => {
      throw new Error('fetch should not run');
    });
    await assert.rejects(
      () => registerListWebhook(deps, { id: 3, isActive: false, webhookSubscriptionId: '' }),
      /Activate/
    );
    const placeholder = createSharePointDeps({
      ...ENV,
      SPO_WEBHOOK_NOTIFICATION_URL: 'https://<function-app>.azurewebsites.net/api/spoWebhook?code=<function-key>'
    }, async () => {
      throw new Error('fetch should not run');
    });
    await assert.rejects(
      () => registerListWebhook(placeholder, {
        id: 3,
        isActive: true,
        webhookSubscriptionId: '',
        targetListUrl: 'https://contoso.sharepoint.com/sites/ops/Lists/Requests',
        targetListGuid: LIST_ID
      }),
      /SPO_WEBHOOK_NOTIFICATION_URL/
    );
  });

  it('creates a list subscription and writes the id back with If-Match', async () => {
    const calls = [];
    const fetchImpl = async (url, options = {}) => {
      calls.push({ url, options });
      if (url.indexOf('oauth2') >= 0) {
        return jsonResponse(200, { access_token: 'token', expires_in: 3600 });
      }
      if (url.endsWith('/subscriptions')) {
        const body = JSON.parse(options.body);
        assert.equal(body.resource, `https://contoso.sharepoint.com/sites/ops/_api/web/lists('${LIST_ID}')`);
        assert.equal(body.notificationUrl, ENV.SPO_WEBHOOK_NOTIFICATION_URL);
        assert.equal(body.clientState, ENV.SHAREPOINT_WEBHOOK_CLIENT_STATE);
        assert.ok(body.expirationDateTime);
        return jsonResponse(201, { Id: 'sub-new' });
      }
      assert.equal(options.headers['X-HTTP-Method'], 'MERGE');
      assert.equal(options.headers['IF-MATCH'], '"9"');
      assert.deepEqual(JSON.parse(options.body), { WebhookSubscriptionId: 'sub-new' });
      return jsonResponse(204, '');
    };
    const result = await registerListWebhook(createSharePointDeps(ENV, fetchImpl), {
      id: 9,
      etag: '"9"',
      isActive: true,
      webhookSubscriptionId: '',
      targetListUrl: 'https://contoso.sharepoint.com/sites/ops/Lists/Requests/AllItems.aspx',
      targetListGuid: LIST_ID
    });
    assert.deepEqual(result, { subscriptionId: 'sub-new', alreadyRegistered: false });
    assert.equal(calls.length, 3);
  });
});

describe('SharePoint URL helpers', () => {
  it('builds a recent-items query and treats empty numbers as blank', () => {
    const url = buildRecentItemsUrl(
      'https://contoso.sharepoint.com/sites/ops/',
      `{${LIST_ID}}`,
      'RequestNumber'
    );
    assert.equal(
      url,
      `https://contoso.sharepoint.com/sites/ops/_api/web/lists(guid'${LIST_ID}')/items?$select=Id,RequestNumber&$orderby=Id desc&$top=200`
    );
    assert.equal(isBlankNumber(null), true);
    assert.equal(isBlankNumber('  '), true);
    assert.equal(isBlankNumber('REQ-0001'), false);
    assert.equal(normalizeGuid(`{${LIST_ID.toUpperCase()}}`), LIST_ID);
  });

  it('derives the web from a list URL and a library AllItems URL', () => {
    assert.deepEqual(
      deriveWebFromListUrl('https://contoso.sharepoint.com/sites/ops/Lists/Requests/AllItems.aspx'),
      {
        webAbsoluteUrl: 'https://contoso.sharepoint.com/sites/ops',
        serverRelativeUrl: '/sites/ops/Lists/Requests'
      }
    );
    assert.deepEqual(
      deriveWebFromListUrl('https://contoso.sharepoint.com/sites/ops/Shared Documents/Forms/AllItems.aspx'),
      {
        webAbsoluteUrl: 'https://contoso.sharepoint.com/sites/ops',
        serverRelativeUrl: '/sites/ops/Shared Documents'
      }
    );
    assert.equal(
      listIdFromResource(`https://contoso.sharepoint.com/sites/ops/_api/web/lists('${LIST_ID.toUpperCase()}')`),
      LIST_ID
    );
  });
});
