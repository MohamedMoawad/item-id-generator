'use strict';

const { app } = require('@azure/functions');
const { generateItemId } = require('../rules/idRules');
const { notifyWebhook } = require('../webhook/notifyWebhook');

/**
 * HTTP POST /api/GenerateItemId
 *
 * Body: { listName, condition, itemId, ...optional metadata }
 * Success response: HTTP 200 with JSON body { id }.
 * The v4 model serializes `jsonBody` to that HTTP body.
 */
async function generateItemIdHandler(request, context, dependencies = {}) {
  const parsed = await readJsonObject(request);
  if (parsed.error) {
    return json(400, { error: parsed.error });
  }

  const item = parseGenerateRequest(parsed.value);
  if (item.error) {
    return json(400, { error: item.error });
  }

  const generate = dependencies.generateItemId || generateItemId;
  const id = generate(item.value);
  const webhookPayload = {
    id,
    itemId: item.value.itemId,
    condition: item.value.condition,
    listName: item.value.listName
  };

  try {
    await notifyWebhook(webhookPayload, dependencies.webhook);
  } catch (error) {
    const reason = error instanceof Error ? error.message : 'unknown error';
    // Generation already succeeded. Do not fail the caller for an optional webhook.
    if (context && typeof context.error === 'function') {
      context.error(`Optional webhook failed. The generated id is still returned. ${reason}`);
    }
  }

  return json(200, { id });
}

function json(status, payload) {
  return { status, jsonBody: payload };
}

async function readJsonObject(request) {
  try {
    const value = await request.json();
    if (typeof value !== 'object' || value === null || Array.isArray(value)) {
      return { error: 'Request body must be a JSON object.' };
    }
    return { value };
  } catch {
    return { error: 'Request body must be JSON.' };
  }
}

function parseGenerateRequest(body) {
  if (typeof body.listName !== 'string' || body.listName.trim() === '') {
    return { error: 'listName is required.' };
  }

  const itemId = readItemId(body.itemId);
  if (itemId === undefined) {
    return { error: 'itemId must be a positive integer.' };
  }

  const condition = typeof body.condition === 'string' ? body.condition.trim() : '';
  const metadata = {};
  for (const [key, value] of Object.entries(body)) {
    if (key !== 'listName' && key !== 'condition' && key !== 'itemId') {
      metadata[key] = value;
    }
  }

  return {
    value: {
      listName: body.listName.trim(),
      condition,
      itemId,
      metadata
    }
  };
}

function readItemId(value) {
  if (typeof value === 'number' && Number.isInteger(value) && value > 0) {
    return value;
  }
  if (typeof value === 'string' && /^[1-9]\d*$/.test(value.trim())) {
    const parsed = Number(value.trim());
    if (Number.isSafeInteger(parsed) && parsed > 0) {
      return parsed;
    }
  }
  return undefined;
}

app.http('GenerateItemId', {
  methods: ['POST'],
  authLevel: 'function',
  handler: generateItemIdHandler
});

module.exports = {
  generateItemIdHandler
};
