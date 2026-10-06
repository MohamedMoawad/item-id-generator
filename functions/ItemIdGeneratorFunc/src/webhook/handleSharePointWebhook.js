'use strict';

const crypto = require('crypto');
const { listIdFromResource } = require('../sharepoint/sharePointUrls');

/**
 * SharePoint list webhook endpoint.
 * Subscription handshake: POST ?validationtoken=... must echo the token as text/plain.
 * Change notification: enqueue one queue message per list and return 200 immediately.
 * SharePoint does not include the new item id. The queue worker finds blank number columns.
 */
async function handleSharePointWebhook(request, context, dependencies = {}) {
  const validationToken = readValidationToken(request);
  if (validationToken) {
    return {
      status: 200,
      headers: { 'Content-Type': 'text/plain; charset=utf-8' },
      body: validationToken
    };
  }

  const env = dependencies.env || process.env;
  const configuredState = typeof env.SHAREPOINT_WEBHOOK_CLIENT_STATE === 'string'
    ? env.SHAREPOINT_WEBHOOK_CLIENT_STATE
    : '';
  if (!isUsableSecret(configuredState)) {
    return {
      status: 500,
      jsonBody: { error: 'SHAREPOINT_WEBHOOK_CLIENT_STATE is not configured.' }
    };
  }

  const parsed = await readNotificationBody(request);
  if (parsed.error) {
    return { status: 400, jsonBody: { error: parsed.error } };
  }

  const messages = [];
  for (const entry of parsed.value) {
    if (!clientStateMatches(configuredState, entry && entry.clientState)) {
      return { status: 403, jsonBody: { error: 'Webhook clientState does not match.' } };
    }
    const listId = listIdFromResource(entry && entry.resource);
    if (!listId) {
      return { status: 400, jsonBody: { error: 'Webhook resource must be a list GUID.' } };
    }
    const siteUrl = typeof entry.siteUrl === 'string' ? entry.siteUrl.trim() : '';
    if (!siteUrl) {
      return { status: 400, jsonBody: { error: 'Webhook notification is missing siteUrl.' } };
    }
    messages.push({
      siteUrl,
      listId,
      subscriptionId: typeof entry.subscriptionId === 'string' ? entry.subscriptionId : ''
    });
  }

  if (messages.length > 0) {
    const enqueue = dependencies.enqueue || defaultEnqueue(context);
    await enqueue(messages);
  }

  if (context && typeof context.log === 'function') {
    context.log(`Accepted ${messages.length} SharePoint webhook notification(s).`);
  }
  return { status: 200, jsonBody: { accepted: messages.length } };
}

function defaultEnqueue(context) {
  return async () => {
    throw new Error('Webhook enqueue is not configured.');
  };
}

function readValidationToken(request) {
  if (request && request.query && typeof request.query.get === 'function') {
    const fromQuery = request.query.get('validationtoken') || request.query.get('validationToken');
    if (fromQuery) {
      return fromQuery;
    }
  }
  try {
    const url = new URL(request.url, 'https://localhost');
    return url.searchParams.get('validationtoken') || url.searchParams.get('validationToken') || '';
  } catch {
    return '';
  }
}

async function readNotificationBody(request) {
  try {
    const body = await request.json();
    if (!body || typeof body !== 'object' || !Array.isArray(body.value)) {
      return { error: 'Webhook body must be a JSON object with a value array.' };
    }
    return { value: body.value };
  } catch {
    return { error: 'Webhook body must be JSON.' };
  }
}

function isUsableSecret(value) {
  const text = value.trim();
  return text.length > 0 && text.indexOf('<') === -1;
}

function clientStateMatches(expected, received) {
  const left = Buffer.from(expected);
  const right = Buffer.from(typeof received === 'string' ? received : '');
  if (left.length !== right.length) {
    return false;
  }
  return crypto.timingSafeEqual(left, right);
}

module.exports = {
  handleSharePointWebhook
};
