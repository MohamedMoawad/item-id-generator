'use strict';

/**
 * POSTs the generated id to WEBHOOK_URL when that setting is non-empty.
 * Callers should treat failures as non-fatal: the id was already generated.
 */
async function notifyWebhook(payload, options = {}) {
  const env = options.env || process.env;
  const fetchImpl = options.fetchImpl || globalThis.fetch;
  const webhookUrl = typeof env.WEBHOOK_URL === 'string' ? env.WEBHOOK_URL.trim() : '';

  if (!webhookUrl) {
    return { skipped: true };
  }

  const body = {
    id: payload.id,
    itemId: payload.itemId,
    condition: payload.condition,
    listName: payload.listName
  };

  let response;
  try {
    response = await fetchImpl(webhookUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
      signal: AbortSignal.timeout(10000)
    });
  } catch (error) {
    const reason = error instanceof Error ? error.message : 'unknown error';
    throw new Error(`Webhook request failed: ${reason}`);
  }

  if (!response || response.ok !== true) {
    const status = response && response.status ? response.status : 'no-status';
    throw new Error(`Webhook returned HTTP ${status}.`);
  }

  return { skipped: false, status: response.status };
}

module.exports = {
  notifyWebhook
};
