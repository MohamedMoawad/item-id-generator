'use strict';

const CONFIG_LIST_TITLE = 'RequestNumberConfig';

const CONFIG_FIELDS = [
  'Id',
  'Title',
  'TargetListUrl',
  'TargetListGuid',
  'NumberColumnInternalName',
  'Formula',
  'CurrentCount',
  'ResetPeriod',
  'LastResetDate',
  'IsActive',
  'PadLength',
  'WebhookSubscriptionId'
];

function normalizeGuid(value) {
  const text = String(value ?? '').trim().replace(/[{}]/g, '').toLowerCase();
  if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/.test(text)) {
    return '';
  }
  return text;
}

function listIdFromResource(resource) {
  const direct = normalizeGuid(resource);
  if (direct) {
    return direct;
  }
  const match = String(resource ?? '').match(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i);
  return match ? normalizeGuid(match[0]) : '';
}

function assertInternalName(name) {
  if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(String(name || ''))) {
    throw new Error(`Invalid column internal name "${name}".`);
  }
}

function assertHttpsUrl(value, label) {
  let url;
  try {
    url = new URL(value);
  } catch {
    throw new Error(`${label} is not a URL.`);
  }
  if (url.protocol !== 'https:') {
    throw new Error(`${label} must use https.`);
  }
  return url;
}

function siteRoot(siteUrl) {
  return assertHttpsUrl(siteUrl, 'SharePoint site URL').toString().replace(/\/+$/, '');
}

function readSiteCollectionUrl(payload, fallbackWebUrl) {
  const fallback = siteRoot(fallbackWebUrl);
  const record = payload && typeof payload === 'object' ? payload : {};
  const nested = record.d && typeof record.d === 'object' ? record.d : {};
  const url = record.Url || record.url || nested.Url || nested.url || '';
  if (typeof url === 'string' && url.trim()) {
    return siteRoot(url);
  }
  return fallback;
}

function escapeODataString(value) {
  return String(value).replace(/'/g, "''");
}

function buildConfigItemsUrl(siteUrl, listTitle) {
  const select = CONFIG_FIELDS.join(',');
  return `${siteRoot(siteUrl)}/_api/web/lists/getbytitle('${escapeODataString(listTitle)}')/items?$select=${select}&$top=500`;
}

function buildConfigItemUrl(siteUrl, listTitle, itemId) {
  return `${siteRoot(siteUrl)}/_api/web/lists/getbytitle('${escapeODataString(listTitle)}')/items(${itemId})`;
}

function buildTargetItemUrl(siteUrl, listId, itemId) {
  const guid = normalizeGuid(listId);
  if (!guid) {
    throw new Error('Target list id must be a GUID.');
  }
  return `${siteRoot(siteUrl)}/_api/web/lists(guid'${guid}')/items(${itemId})`;
}

function buildRecentItemsUrl(siteUrl, listId, fieldInternalName) {
  assertInternalName(fieldInternalName);
  const guid = normalizeGuid(listId);
  if (!guid) {
    throw new Error('Target list id must be a GUID.');
  }
  return `${siteRoot(siteUrl)}/_api/web/lists(guid'${guid}')/items?$select=Id,${fieldInternalName}&$orderby=Id desc&$top=200`;
}

function isBlankNumber(value) {
  return value === undefined || value === null || String(value).trim() === '';
}

function readEtag(item) {
  if (!item || typeof item !== 'object') {
    return '';
  }
  return item['odata.etag'] || item['@odata.etag'] || '';
}

function isActiveValue(value) {
  return value === true || value === 1 || value === '1' || value === 'true' || value === 'Yes';
}

function deriveWebFromListUrl(listUrl) {
  const url = assertHttpsUrl(listUrl, 'Target list URL');
  let path = decodeURIComponent(url.pathname);
  path = path.replace(/\/Forms\/AllItems\.aspx$/i, '');
  path = path.replace(/\/AllItems\.aspx$/i, '');
  path = path.replace(/\/+$/, '');
  const listsMatch = path.match(/^(.*)\/lists\/[^/]+$/i);
  const webPath = listsMatch ? (listsMatch[1] || '') : path.slice(0, path.lastIndexOf('/'));
  return {
    webAbsoluteUrl: `${url.origin}${webPath || ''}`,
    serverRelativeUrl: path
  };
}

function normalizeConfigItem(raw) {
  if (!raw || typeof raw !== 'object') {
    return undefined;
  }
  const targetListUrl = typeof raw.TargetListUrl === 'string' ? raw.TargetListUrl.trim() : '';
  let targetSiteUrl = '';
  if (targetListUrl) {
    try {
      targetSiteUrl = deriveWebFromListUrl(targetListUrl).webAbsoluteUrl;
    } catch {
      targetSiteUrl = '';
    }
  }
  const padLength = Number(raw.PadLength);
  return {
    id: raw.Id,
    etag: readEtag(raw),
    title: raw.Title || '',
    targetListUrl,
    targetSiteUrl,
    targetListGuid: normalizeGuid(raw.TargetListGuid),
    numberColumnInternalName: raw.NumberColumnInternalName || '',
    currentCount: Number(raw.CurrentCount) || 0,
    formula: raw.Formula || '',
    resetPeriod: raw.ResetPeriod || 'None',
    lastResetDate: raw.LastResetDate || '',
    isActive: isActiveValue(raw.IsActive),
    padLength: Number.isInteger(padLength) && padLength > 0 ? padLength : 0,
    webhookSubscriptionId: raw.WebhookSubscriptionId || ''
  };
}

module.exports = {
  CONFIG_FIELDS,
  CONFIG_LIST_TITLE,
  assertHttpsUrl,
  assertInternalName,
  buildConfigItemUrl,
  buildConfigItemsUrl,
  buildRecentItemsUrl,
  buildTargetItemUrl,
  deriveWebFromListUrl,
  escapeODataString,
  isActiveValue,
  isBlankNumber,
  listIdFromResource,
  normalizeConfigItem,
  normalizeGuid,
  readEtag,
  readSiteCollectionUrl
};
