'use strict';

const {
  createSharePointDeps,
  getConfigItemById,
  registerListWebhook
} = require('../sharepoint/sharePointClient');

/**
 * App-only registration. The browser cannot finish a SharePoint webhook
 * handshake by itself in every tenant, and the function already holds the
 * notification URL and clientState. SPFx posts the config row id after save.
 */
async function handleRegisterWebhook(request, context, dependencies = {}) {
  const parsed = await readBody(request);
  if (parsed.error) {
    return { status: 400, jsonBody: { error: parsed.error } };
  }

  const configItemId = parsed.value.configItemId;
  if (typeof configItemId !== 'number' || !Number.isInteger(configItemId) || configItemId <= 0) {
    return { status: 400, jsonBody: { error: 'configItemId must be a positive integer.' } };
  }

  const configSiteUrl = typeof parsed.value.configSiteUrl === 'string' ? parsed.value.configSiteUrl.trim() : '';
  if (!configSiteUrl || configSiteUrl.indexOf('<') >= 0) {
    return { status: 400, jsonBody: { error: 'configSiteUrl must be the https URL of the site collection.' } };
  }

  const env = { ...(dependencies.env || process.env) };
  env.NUMBERING_CONFIG_SITE_URL = configSiteUrl;

  try {
    const deps = dependencies.sharePointDeps || createSharePointDeps(env, dependencies.fetchImpl || globalThis.fetch);
    const config = await getConfigItemById(deps, configItemId);
    if (!config) {
      return { status: 404, jsonBody: { error: 'Config row was not found.' } };
    }
    const result = await registerListWebhook(deps, config);
    if (context && typeof context.log === 'function') {
      context.log(`Webhook ${result.subscriptionId} for config item ${configItemId}.`);
    }
    return { status: 200, jsonBody: result };
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Webhook registration failed.';
    return { status: 502, jsonBody: { error: message } };
  }
}

async function readBody(request) {
  try {
    const value = await request.json();
    if (!value || typeof value !== 'object' || Array.isArray(value)) {
      return { error: 'Request body must be a JSON object.' };
    }
    return { value };
  } catch {
    return { error: 'Request body must be JSON.' };
  }
}

module.exports = {
  handleRegisterWebhook
};
