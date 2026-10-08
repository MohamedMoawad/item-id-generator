import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import {
  buildAllPropertiesUrl,
  buildPropertyBagMerge,
  buildSetStorageEntity,
  buildStorageEntityUrl,
  buildTenantSettingsUrl,
  buildCreateBody,
  buildCreateFieldUrl,
  buildCreateItemUrl,
  buildCreateListBody,
  buildCreateListUrl,
  buildFieldsUrl,
  buildGetListUrl,
  buildListItemsUrl,
  buildListPermissionsUrl,
  buildListUrl,
  buildRegisterWebhookBody,
  buildUpdateBody,
  buildUpdateItemUrl,
  deriveWebFromListUrl,
  findConfigForList,
  hasPermissionFlag,
  isPlaceholderSetting,
  IConfigDraft,
  ISharePointConfigItem,
  mapConfigRow,
  readCorporateCatalogUrl,
  readStorageEntityValue,
  PERMISSION_EDIT_LIST_ITEMS,
  REGISTER_WEBHOOK_PROPERTY,
  sharePointError
} from './configContract';
import { buildCreateFieldBody, CONFIG_FIELD_DEFINITIONS } from './requestNumberFields';

export class ConfigListRequestError extends Error {
  public readonly status: number;

  public constructor(message: string, status: number) {
    super(message);
    this.name = 'ConfigListRequestError';
    this.status = status;
  }
}

export interface IEnsureListResult {
  created: boolean;
  fieldsAdded: string[];
}

export interface IResolvedList {
  id: string;
  title: string;
}

const jsonHeaders: { [key: string]: string } = {
  Accept: 'application/json;odata=nometadata',
  'Content-Type': 'application/json;odata=nometadata'
};

// SPHttpClient v1 sends OData-Version 4.0. CreateFieldAsXml rejects the v3
// __metadata property on SP.XmlSchemaFieldCreationInformation (HTTP 400).
const fieldHeaders: { [key: string]: string } = {
  Accept: 'application/json;odata.metadata=none',
  'Content-Type': 'application/json;charset=utf-8'
};

export async function loadConfigRows(
  spHttpClient: SPHttpClient,
  siteAbsoluteUrl: string,
  listTitle: string
): Promise<IConfigDraft[]> {
  const response = await spHttpClient.get(
    buildListItemsUrl(siteAbsoluteUrl, listTitle),
    SPHttpClient.configurations.v1,
    { headers: { Accept: 'application/json;odata=minimalmetadata' } }
  );
  const text = await response.text();
  if (!response.ok) {
    throw new ConfigListRequestError(sharePointError(response.status, text, listTitle), response.status);
  }
  const payload = parseJson<{ value?: ISharePointConfigItem[] }>(text);
  return (payload.value || []).map(mapConfigRow);
}

export async function ensureConfigList(
  spHttpClient: SPHttpClient,
  siteAbsoluteUrl: string,
  listTitle: string
): Promise<IEnsureListResult> {
  const existing = await spHttpClient.get(
    buildListUrl(siteAbsoluteUrl, listTitle),
    SPHttpClient.configurations.v1,
    { headers: { Accept: 'application/json;odata=nometadata' } }
  );
  let created = false;
  if (existing.status === 404) {
    const createResponse = await spHttpClient.post(
      buildCreateListUrl(siteAbsoluteUrl),
      SPHttpClient.configurations.v1,
      {
        headers: jsonHeaders,
        body: JSON.stringify(buildCreateListBody(listTitle))
      }
    );
    if (!createResponse.ok) {
      throw new ConfigListRequestError(
        sharePointError(createResponse.status, await createResponse.text(), listTitle),
        createResponse.status
      );
    }
    created = true;
  } else if (!existing.ok) {
    throw new ConfigListRequestError(
      sharePointError(existing.status, await existing.text(), listTitle),
      existing.status
    );
  }

  const fieldsResponse = await spHttpClient.get(
    buildFieldsUrl(siteAbsoluteUrl, listTitle),
    SPHttpClient.configurations.v1,
    { headers: { Accept: 'application/json;odata=nometadata' } }
  );
  const fieldsText = await fieldsResponse.text();
  if (!fieldsResponse.ok) {
    throw new ConfigListRequestError(sharePointError(fieldsResponse.status, fieldsText, listTitle), fieldsResponse.status);
  }
  const fieldPayload = parseJson<{
    value?: { InternalName?: string }[];
    d?: { results?: { InternalName?: string }[] };
  }>(fieldsText);
  const present = new Set(readInternalNames(fieldPayload));
  const fieldsAdded: string[] = [];
  for (const field of CONFIG_FIELD_DEFINITIONS) {
    if (present.has(field.internalName)) {
      continue;
    }
    const addResponse = await spHttpClient.post(
      buildCreateFieldUrl(siteAbsoluteUrl, listTitle),
      SPHttpClient.configurations.v1,
      {
        headers: fieldHeaders,
        body: JSON.stringify(buildCreateFieldBody(field.schemaXml))
      }
    );
    if (!addResponse.ok) {
      throw new ConfigListRequestError(
        sharePointError(addResponse.status, await addResponse.text(), listTitle),
        addResponse.status
      );
    }
    fieldsAdded.push(field.internalName);
  }
  return { created, fieldsAdded };
}

export async function resolveTargetList(spHttpClient: SPHttpClient, listUrl: string): Promise<IResolvedList> {
  const response = await spHttpClient.get(
    buildGetListUrl(listUrl),
    SPHttpClient.configurations.v1,
    { headers: { Accept: 'application/json;odata=nometadata' } }
  );
  const text = await response.text();
  if (response.status === 404) {
    throw new ConfigListRequestError('That list URL was not found. Check the address and your permissions on that site.', 404);
  }
  if (!response.ok) {
    throw new ConfigListRequestError(sharePointError(response.status, text, 'target list'), response.status);
  }
  const payload = parseJson<{ Id?: string; Title?: string; d?: { Id?: string; Title?: string } }>(text);
  const id = (payload.Id || (payload.d && payload.d.Id) || '').replace(/[{}]/g, '').toLowerCase();
  const title = payload.Title || (payload.d && payload.d.Title) || '';
  if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/.test(id)) {
    throw new Error('SharePoint did not return a list GUID for that URL.');
  }
  return { id, title };
}

export async function saveConfigRow(
  spHttpClient: SPHttpClient,
  siteAbsoluteUrl: string,
  listTitle: string,
  draft: IConfigDraft
): Promise<number> {
  const itemId = draft.id;
  const creating = itemId === undefined;
  const url = creating
    ? buildCreateItemUrl(siteAbsoluteUrl, listTitle)
    : buildUpdateItemUrl(siteAbsoluteUrl, listTitle, itemId);
  const response: SPHttpClientResponse = await spHttpClient.post(url, SPHttpClient.configurations.v1, {
    headers: creating
      ? jsonHeaders
      : {
        ...jsonHeaders,
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE'
      },
    body: JSON.stringify(creating ? buildCreateBody(draft) : buildUpdateBody(draft))
  });

  if (!response.ok) {
    throw new ConfigListRequestError(sharePointError(response.status, await response.text(), listTitle), response.status);
  }
  if (!creating) {
    return itemId;
  }
  const payload = parseJson<{ Id?: number; d?: { Id?: number } }>(await response.text());
  const createdId = payload.Id || (payload.d && payload.d.Id);
  if (!createdId) {
    throw new Error('SharePoint did not return the new config item id.');
  }
  return createdId;
}

export async function ensureTargetNumberColumn(
  spHttpClient: SPHttpClient,
  listUrl: string,
  listGuid: string,
  internalName: string
): Promise<void> {
  const name = internalName.trim();
  if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) {
    throw new Error('Number column internal name must look like RequestNumber.');
  }
  const web = deriveWebFromListUrl(listUrl).webAbsoluteUrl;
  const guid = listGuid.replace(/[{}]/g, '');
  const fieldUrl = `${web}/_api/web/lists(guid'${guid}')/fields/getbyinternalnameorstaticname('${name}')`;
  const existing = await spHttpClient.get(fieldUrl, SPHttpClient.configurations.v1, {
    headers: { Accept: 'application/json;odata=nometadata' }
  });
  if (existing.ok) {
    return;
  }
  const schema = `<Field Type="Text" Name="${name}" StaticName="${name}" DisplayName="${name}" />`;
  const response = await spHttpClient.post(
    `${web}/_api/web/lists(guid'${guid}')/fields/CreateFieldAsXml`,
    SPHttpClient.configurations.v1,
    {
      headers: fieldHeaders,
      body: JSON.stringify(buildCreateFieldBody(schema))
    }
  );
  if (response.ok) {
    return;
  }
  const detail = await response.text();
  if (response.status === 400 && /already exists|duplicate/i.test(detail)) {
    return;
  }
  throw new ConfigListRequestError(
    sharePointError(response.status, detail, name),
    response.status
  );
}

export async function connectConfiguredList(
  spHttpClient: SPHttpClient,
  siteAbsoluteUrl: string,
  listTitle: string,
  listGuid: string
): Promise<boolean> {
  const webhookUrl = await readRegisterWebhookUrl(spHttpClient, siteAbsoluteUrl);
  if (isPlaceholderSetting(webhookUrl)) {
    return false;
  }
  const rows = await loadConfigRows(spHttpClient, siteAbsoluteUrl, listTitle);
  const row = findConfigForList(rows, listGuid);
  if (!row || !row.isActive || row.id === undefined) {
    return false;
  }
  if (row.webhookSubscriptionId) {
    return true;
  }
  await callRegisterWebhook(webhookUrl, siteAbsoluteUrl, row.id);
  return true;
}

export async function callRegisterWebhook(
  registerWebhookUrl: string,
  configSiteUrl: string,
  configItemId: number
): Promise<{ subscriptionId: string; alreadyRegistered: boolean }> {
  let response: Response;
  try {
    response = await fetch(registerWebhookUrl, {
      method: 'POST',
      headers: {
        Accept: 'application/json',
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(buildRegisterWebhookBody(configSiteUrl, configItemId))
    });
  } catch (error) {
    const detail = error instanceof Error ? error.message : 'network error';
    throw new Error(`The browser could not reach the function (${detail}). On the Function App open CORS and add the SharePoint origin, then save again.`);
  }
  const text = await response.text();
  const payload = text
    ? parseJson<{ subscriptionId?: string; alreadyRegistered?: boolean; error?: string }>(text)
    : {};
  if (!response.ok) {
    const fallback = response.status === 401
      ? 'The function rejected the key. Copy RegisterWebhook → Get function URL, including ?code=.'
      : `RegisterWebhook returned HTTP ${response.status}.`;
    throw new Error(payload.error || fallback);
  }
  if (!payload.subscriptionId) {
    throw new Error('RegisterWebhook did not return a subscription id.');
  }
  return {
    subscriptionId: payload.subscriptionId,
    alreadyRegistered: payload.alreadyRegistered === true
  };
}

export async function canEditConfigList(
  spHttpClient: SPHttpClient,
  siteAbsoluteUrl: string,
  listTitle: string
): Promise<boolean> {
  const response = await spHttpClient.get(
    buildListPermissionsUrl(siteAbsoluteUrl, listTitle),
    SPHttpClient.configurations.v1,
    { headers: { Accept: 'application/json;odata=nometadata' } }
  );
  if (!response.ok) {
    return false;
  }
  const payload = parseJson<{ Low?: string | number; d?: { Low?: string | number } }>(await response.text());
  return hasPermissionFlag(payload.Low ?? (payload.d && payload.d.Low), PERMISSION_EDIT_LIST_ITEMS);
}

export async function readRegisterWebhookUrl(spHttpClient: SPHttpClient, siteAbsoluteUrl: string): Promise<string> {
  const organization = await readOrganizationWebhookUrl(spHttpClient, siteAbsoluteUrl);
  if (organization) {
    return organization;
  }
  const response = await spHttpClient.get(
    buildAllPropertiesUrl(siteAbsoluteUrl),
    SPHttpClient.configurations.v1,
    { headers: { Accept: 'application/json;odata=nometadata' } }
  );
  if (!response.ok) {
    return '';
  }
  const payload = parseJson<Record<string, unknown>>(await response.text());
  const direct = payload[REGISTER_WEBHOOK_PROPERTY];
  if (typeof direct === 'string') {
    return direct;
  }
  const nested = payload.d;
  if (nested && typeof nested === 'object' && typeof (nested as Record<string, unknown>)[REGISTER_WEBHOOK_PROPERTY] === 'string') {
    return (nested as Record<string, unknown>)[REGISTER_WEBHOOK_PROPERTY] as string;
  }
  return '';
}

export async function saveRegisterWebhookUrl(
  spHttpClient: SPHttpClient,
  siteAbsoluteUrl: string,
  registerWebhookUrl: string
): Promise<{ organization: boolean }> {
  const organization = await trySaveOrganizationWebhookUrl(spHttpClient, siteAbsoluteUrl, registerWebhookUrl.trim());
  const merge = buildPropertyBagMerge(REGISTER_WEBHOOK_PROPERTY, registerWebhookUrl.trim());
  const response = await spHttpClient.post(
    buildAllPropertiesUrl(siteAbsoluteUrl),
    SPHttpClient.configurations.v1,
    {
      headers: merge.headers,
      body: merge.body
    }
  );
  if (!response.ok) {
    throw new ConfigListRequestError(
      sharePointError(response.status, await response.text(), 'site properties'),
      response.status
    );
  }
  return { organization };
}

async function readOrganizationWebhookUrl(spHttpClient: SPHttpClient, siteAbsoluteUrl: string): Promise<string> {
  try {
    const response = await spHttpClient.get(
      buildStorageEntityUrl(siteAbsoluteUrl),
      SPHttpClient.configurations.v1,
      { headers: { Accept: 'application/json;odata=verbose' } }
    );
    if (!response.ok) {
      return '';
    }
    const value = readStorageEntityValue(parseJson<unknown>(await response.text()));
    return value && !isPlaceholderSetting(value) ? value : '';
  } catch {
    return '';
  }
}

async function trySaveOrganizationWebhookUrl(
  spHttpClient: SPHttpClient,
  siteAbsoluteUrl: string,
  registerWebhookUrl: string
): Promise<boolean> {
  try {
    const settings = await spHttpClient.get(
      buildTenantSettingsUrl(siteAbsoluteUrl),
      SPHttpClient.configurations.v1,
      { headers: { Accept: 'application/json;odata=verbose' } }
    );
    if (!settings.ok) {
      return false;
    }
    const catalog = readCorporateCatalogUrl(parseJson<unknown>(await settings.text()));
    if (!catalog) {
      return false;
    }
    const request = buildSetStorageEntity(catalog, registerWebhookUrl);
    const response = await spHttpClient.post(request.url, SPHttpClient.configurations.v1, {
      headers: request.headers,
      body: request.body
    });
    return response.ok;
  } catch {
    return false;
  }
}

function readInternalNames(payload: { value?: { InternalName?: string }[]; d?: { results?: { InternalName?: string }[] } }): string[] {
  const rows = payload.value || (payload.d && payload.d.results) || [];
  return rows
    .map((row) => row.InternalName || '')
    .filter((name) => name.length > 0);
}

function parseJson<T>(text: string): T {
  try {
    return JSON.parse(text) as T;
  } catch {
    throw new Error('SharePoint returned a response that was not JSON.');
  }
}
