'use strict';

const {
  assertHttpsUrl,
  assertInternalName,
  buildConfigItemUrl,
  buildConfigItemsUrl,
  buildRecentItemsUrl,
  buildTargetItemUrl,
  deriveWebFromListUrl,
  isBlankNumber,
  normalizeConfigItem,
  normalizeGuid
} = require('./sharePointUrls');

const DEFAULT_CONFIG_LIST_TITLE = 'RequestNumberConfig';

function configListTitle(env) {
  const title = (env.NUMBERING_CONFIG_LIST_TITLE || DEFAULT_CONFIG_LIST_TITLE).trim();
  if (!title || /[\r\n]/.test(title)) {
    throw new Error('NUMBERING_CONFIG_LIST_TITLE is invalid.');
  }
  return title;
}

function configSiteUrl(env) {
  const siteUrl = (env.NUMBERING_CONFIG_SITE_URL || '').trim();
  if (!siteUrl || siteUrl.indexOf('<') >= 0) {
    throw new Error('Set NUMBERING_CONFIG_SITE_URL to the site collection that contains RequestNumberConfig.');
  }
  return assertHttpsUrl(siteUrl, 'NUMBERING_CONFIG_SITE_URL').toString().replace(/\/+$/, '');
}

function assertAppIdentity(env) {
  const tenantId = (env.SHAREPOINT_TENANT_ID || '').trim();
  const clientId = (env.SHAREPOINT_CLIENT_ID || '').trim();
  const clientSecret = env.SHAREPOINT_CLIENT_SECRET || '';
  if (!normalizeGuid(tenantId) || !normalizeGuid(clientId) || !clientSecret || clientSecret.indexOf('<') >= 0) {
    throw new Error('Set SHAREPOINT_TENANT_ID, SHAREPOINT_CLIENT_ID, and SHAREPOINT_CLIENT_SECRET. Do not commit the secret.');
  }
  return { tenantId, clientId, clientSecret };
}

async function readBody(response) {
  try {
    return await response.text();
  } catch {
    return '';
  }
}

function httpError(response, body, action) {
  const detail = String(body || '').trim().slice(0, 300);
  return new Error(`SharePoint ${action} failed (HTTP ${response.status}). ${detail}`);
}

async function getAccessToken(deps, siteUrl) {
  const identity = assertAppIdentity(deps.env);
  const host = assertHttpsUrl(siteUrl, 'SharePoint site URL').host;
  const cache = deps.tokenCache;
  const now = Date.now();
  if (cache.token && cache.host === host && cache.expiresAt > now + 60000) {
    return cache.token;
  }

  const body = new URLSearchParams({
    client_id: identity.clientId,
    client_secret: identity.clientSecret,
    scope: `https://${host}/.default`,
    grant_type: 'client_credentials'
  });
  const response = await deps.fetchImpl(
    `https://login.microsoftonline.com/${identity.tenantId}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body
    }
  );
  const text = await readBody(response);
  if (!response.ok) {
    throw httpError(response, text, 'token request');
  }
  const payload = JSON.parse(text);
  if (!payload.access_token) {
    throw new Error('SharePoint token response did not include an access token.');
  }
  cache.token = payload.access_token;
  cache.host = host;
  cache.expiresAt = now + (Number(payload.expires_in) || 3600) * 1000;
  return cache.token;
}

async function sharePointHeaders(deps, siteUrl, extra) {
  const token = await getAccessToken(deps, siteUrl);
  return {
    Authorization: `Bearer ${token}`,
    Accept: 'application/json;odata=minimalmetadata',
    ...extra
  };
}

function collectionItems(payload) {
  if (Array.isArray(payload)) {
    return payload;
  }
  if (payload && Array.isArray(payload.value)) {
    return payload.value;
  }
  if (payload && payload.d && Array.isArray(payload.d.results)) {
    return payload.d.results;
  }
  return [];
}

async function getEnabledConfig(deps, listId) {
  const guid = normalizeGuid(listId);
  if (!guid) {
    throw new Error('Target list id must be a GUID.');
  }
  const siteUrl = configSiteUrl(deps.env);
  const url = buildConfigItemsUrl(siteUrl, configListTitle(deps.env));
  const response = await deps.fetchImpl(url, {
    method: 'GET',
    headers: await sharePointHeaders(deps, siteUrl)
  });
  const text = await readBody(response);
  if (!response.ok) {
    throw httpError(response, text, 'config list read');
  }
  const matches = collectionItems(JSON.parse(text))
    .map(normalizeConfigItem)
    .filter((item) => item && item.isActive && item.targetListGuid === guid);
  if (matches.length > 1) {
    throw new Error(`More than one active numbering config targets list ${guid}.`);
  }
  return matches[0];
}

async function getConfigItemById(deps, itemId) {
  const siteUrl = configSiteUrl(deps.env);
  const url = buildConfigItemUrl(siteUrl, configListTitle(deps.env), itemId);
  const response = await deps.fetchImpl(url, {
    method: 'GET',
    headers: await sharePointHeaders(deps, siteUrl)
  });
  const text = await readBody(response);
  if (response.status === 404) {
    return undefined;
  }
  if (!response.ok) {
    throw httpError(response, text, 'config item read');
  }
  return normalizeConfigItem(JSON.parse(text));
}

async function mergeConfigItem(deps, itemId, etag, patch) {
  const siteUrl = configSiteUrl(deps.env);
  const url = buildConfigItemUrl(siteUrl, configListTitle(deps.env), itemId);
  const response = await deps.fetchImpl(url, {
    method: 'POST',
    headers: await sharePointHeaders(deps, siteUrl, {
      'Content-Type': 'application/json;odata=nometadata',
      'IF-MATCH': etag,
      'X-HTTP-Method': 'MERGE'
    }),
    body: JSON.stringify({
      CurrentCount: patch.currentCount,
      LastResetDate: patch.lastResetDate
    })
  });
  if (response.status === 412) {
    return { ok: false };
  }
  if (!response.ok) {
    throw httpError(response, await readBody(response), 'config item update');
  }
  return { ok: true };
}

function createConfigStore(deps, config) {
  return {
    async read() {
      return getConfigItemById(deps, config.id);
    },
    async compareAndSwap(etag, patch) {
      return mergeConfigItem(deps, config.id, etag, patch);
    }
  };
}

async function getListItem(deps, siteUrl, listId, itemId, fieldInternalName) {
  assertInternalName(fieldInternalName);
  const url = `${buildTargetItemUrl(siteUrl, listId, itemId)}?$select=Id,${fieldInternalName}`;
  const response = await deps.fetchImpl(url, {
    method: 'GET',
    headers: await sharePointHeaders(deps, siteUrl)
  });
  const text = await readBody(response);
  if (response.status === 404) {
    return undefined;
  }
  if (!response.ok) {
    throw httpError(response, text, 'list item read');
  }
  const raw = JSON.parse(text);
  return {
    id: raw.Id,
    etag: raw['odata.etag'] || raw['@odata.etag'] || '*',
    fields: raw
  };
}

async function listUnnumberedItems(deps, config, explicitItemId) {
  const fieldName = config.numberColumnInternalName;
  assertInternalName(fieldName);
  if (!config.targetSiteUrl) {
    throw new Error('Target list URL is missing or is not an https list URL.');
  }

  if (explicitItemId) {
    const item = await getListItem(deps, config.targetSiteUrl, config.targetListGuid, explicitItemId, fieldName);
    if (!item || !isBlankNumber(item.fields[fieldName])) {
      return [];
    }
    return [{ id: item.id }];
  }

  const url = buildRecentItemsUrl(config.targetSiteUrl, config.targetListGuid, fieldName);
  const response = await deps.fetchImpl(url, {
    method: 'GET',
    headers: await sharePointHeaders(deps, config.targetSiteUrl)
  });
  const text = await readBody(response);
  if (!response.ok) {
    throw httpError(response, text, 'recent items read');
  }
  return collectionItems(JSON.parse(text))
    .filter((item) => isBlankNumber(item[fieldName]))
    .map((item) => ({ id: item.Id }))
    .filter((item) => Number.isInteger(item.id))
    .sort((left, right) => left.id - right.id);
}

async function writeNumberIfBlank(deps, config, itemId, code) {
  const fieldName = config.numberColumnInternalName;
  assertInternalName(fieldName);
  for (let attempt = 1; attempt <= 3; attempt += 1) {
    const item = await getListItem(deps, config.targetSiteUrl, config.targetListGuid, itemId, fieldName);
    if (!item) {
      throw new Error(`List item ${itemId} was not found.`);
    }
    if (!isBlankNumber(item.fields[fieldName])) {
      return { skipped: true };
    }
    const response = await deps.fetchImpl(buildTargetItemUrl(config.targetSiteUrl, config.targetListGuid, itemId), {
      method: 'POST',
      headers: await sharePointHeaders(deps, config.targetSiteUrl, {
        'Content-Type': 'application/json;odata=nometadata',
        'IF-MATCH': item.etag,
        'X-HTTP-Method': 'MERGE'
      }),
      body: JSON.stringify({ [fieldName]: code })
    });
    if (response.status === 412) {
      continue;
    }
    if (!response.ok) {
      throw httpError(response, await readBody(response), 'list item update');
    }
    return { skipped: false };
  }
  throw new Error(`Could not write the request number on item ${itemId}.`);
}

async function mergeConfigFields(deps, itemId, etag, fields) {
  const siteUrl = configSiteUrl(deps.env);
  const url = buildConfigItemUrl(siteUrl, configListTitle(deps.env), itemId);
  const response = await deps.fetchImpl(url, {
    method: 'POST',
    headers: await sharePointHeaders(deps, siteUrl, {
      'Content-Type': 'application/json;odata=nometadata',
      'IF-MATCH': etag || '*',
      'X-HTTP-Method': 'MERGE'
    }),
    body: JSON.stringify(fields)
  });
  if (response.status === 412) {
    return { ok: false };
  }
  if (!response.ok) {
    throw httpError(response, await readBody(response), 'config field update');
  }
  return { ok: true };
}

async function registerListWebhook(deps, config) {
  if (!config || config.isActive !== true) {
    throw new Error('Activate the config row before registering a webhook.');
  }
  if (config.webhookSubscriptionId) {
    return { subscriptionId: config.webhookSubscriptionId, alreadyRegistered: true };
  }
  const notificationUrl = (deps.env.SPO_WEBHOOK_NOTIFICATION_URL || '').trim();
  const clientState = (deps.env.SHAREPOINT_WEBHOOK_CLIENT_STATE || '').trim();
  if (!notificationUrl || notificationUrl.indexOf('<') >= 0) {
    throw new Error('Set SPO_WEBHOOK_NOTIFICATION_URL to the spoWebhook URL, including the function key.');
  }
  if (!clientState || clientState.indexOf('<') >= 0) {
    throw new Error('Set SHAREPOINT_WEBHOOK_CLIENT_STATE before registering a webhook.');
  }
  if (!config.targetListUrl || !config.targetListGuid) {
    throw new Error('The config row needs TargetListUrl and TargetListGuid.');
  }

  const location = deriveWebFromListUrl(config.targetListUrl);
  const expiration = new Date(Date.now() + 170 * 24 * 60 * 60 * 1000).toISOString();
  const response = await deps.fetchImpl(
    `${location.webAbsoluteUrl}/_api/web/lists(guid'${config.targetListGuid}')/subscriptions`,
    {
      method: 'POST',
      headers: await sharePointHeaders(deps, location.webAbsoluteUrl, {
        'Content-Type': 'application/json;odata=nometadata'
      }),
      body: JSON.stringify({
        resource: `${location.webAbsoluteUrl}/_api/web/lists('${config.targetListGuid}')`,
        notificationUrl,
        expirationDateTime: expiration,
        clientState
      })
    }
  );
  const text = await readBody(response);
  if (!response.ok) {
    throw httpError(response, text, 'webhook registration');
  }
  const payload = text ? JSON.parse(text) : {};
  const subscriptionId = payload.Id || payload.id || (payload.d && (payload.d.Id || payload.d.id));
  if (!subscriptionId) {
    throw new Error('SharePoint did not return a webhook subscription id.');
  }
  const saved = await mergeConfigFields(deps, config.id, config.etag, {
    WebhookSubscriptionId: String(subscriptionId)
  });
  if (!saved.ok) {
    throw new Error('The webhook was created but the config row changed. Register it again.');
  }
  return { subscriptionId: String(subscriptionId), alreadyRegistered: false };
}

function createSharePointDeps(env, fetchImpl = globalThis.fetch) {
  return {
    env,
    fetchImpl,
    tokenCache: { token: '', expiresAt: 0, host: '' }
  };
}

module.exports = {
  DEFAULT_CONFIG_LIST_TITLE,
  createConfigStore,
  createSharePointDeps,
  getConfigItemById,
  getEnabledConfig,
  listUnnumberedItems,
  mergeConfigFields,
  mergeConfigItem,
  registerListWebhook,
  writeNumberIfBlank
};
