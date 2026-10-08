'use strict';

const { normalizeGuid } = require('../sharepoint/sharePointUrls');

/**
 * Numbers every still-blank item on one list.
 * The lease keeps a single worker per list GUID. issueNextCode still uses
 * the config row ETag so a lost lease cannot mint the same sequence twice.
 */
async function processNumberingMessage(message, context, dependencies) {
  const work = parseWorkMessage(message);
  const log = (text) => {
    if (context && typeof context.log === 'function') {
      context.log(text);
    }
  };

  const lease = await dependencies.leaseStore.acquire(work.listId);
  if (!lease.acquired) {
    throw new Error(`Numbering lease for list ${work.listId} is held. The queue message will retry.`);
  }

  try {
    if (!work.siteUrl) {
      throw new Error('Queue message siteUrl is required so AutoGenFeatureConfiguration can be read on that site collection.');
    }
    const config = await dependencies.gateway.getEnabledConfig(work.siteUrl, work.listId);
    if (!config) {
      log(`No active numbering config for list ${work.listId}.`);
      return { issued: [] };
    }

    const siteUrl = config.targetSiteUrl || work.siteUrl;
    if (!siteUrl) {
      throw new Error(`Numbering config ${config.id} has no target site URL.`);
    }
    const readyConfig = { ...config, targetSiteUrl: siteUrl };
    const items = (await dependencies.gateway.listUnnumbered(readyConfig, work.itemId))
      .slice()
      .sort((left, right) => left.id - right.id);
    const now = dependencies.now || (() => new Date());
    const issued = [];

    for (const item of items) {
      const next = await dependencies.gateway.reserveNextCode(readyConfig, now());
      const write = await dependencies.gateway.writeNumberIfBlank(readyConfig, item.id, next.code);
      if (!write.skipped) {
        issued.push({ itemId: item.id, code: next.code, sequence: next.sequence });
        log(`Issued ${next.code} to item ${item.id} on list ${work.listId}.`);
      }
    }

    return { issued };
  } finally {
    await lease.release();
  }
}

function parseWorkMessage(message) {
  const body = typeof message === 'string' ? JSON.parse(message) : message;
  if (!body || typeof body !== 'object' || Array.isArray(body)) {
    throw new Error('Queue message must be a JSON object.');
  }
  const listId = normalizeGuid(body.listId);
  if (!listId) {
    throw new Error('Queue message listId must be a GUID.');
  }
  let itemId;
  if (body.itemId !== undefined && body.itemId !== null && body.itemId !== '') {
    itemId = readItemId(body.itemId);
    if (itemId === undefined) {
      throw new Error('Queue message itemId must be a positive integer when it is present.');
    }
  }
  return {
    siteUrl: typeof body.siteUrl === 'string' ? body.siteUrl.trim() : '',
    listId,
    itemId
  };
}

function readItemId(value) {
  if (typeof value === 'number' && Number.isInteger(value) && value > 0) {
    return value;
  }
  if (typeof value === 'string' && /^[1-9]\d*$/.test(value.trim())) {
    const parsed = Number(value.trim());
    if (Number.isSafeInteger(parsed)) {
      return parsed;
    }
  }
  return undefined;
}

module.exports = {
  parseWorkMessage,
  processNumberingMessage
};
