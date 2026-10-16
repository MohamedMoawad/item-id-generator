import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import {
  buildCreateBody,
  buildCreateFieldUrl,
  buildCreateItemUrl,
  buildCreateListBody,
  buildCreateListUrl,
  buildFieldsUrl,
  buildListItemsUrl,
  buildListPermissionsUrl,
  buildListUrl,
  buildUpdateBody,
  buildUpdateItemUrl,
  deriveWebFromListUrl,
  hasPermissionFlag,
  IConfigDraft,
  isExistingColumnError,
  ISharePointConfigItem,
  mapConfigRow,
  PERMISSION_EDIT_LIST_ITEMS,
  sharePointError
} from './contract';
import { formulaColumns, IFormulaColumn } from './formula';
import { buildCreateFieldBody, CONFIG_FIELD_DEFINITIONS } from './fields';

export class ConfigListRequestError extends Error {
  public readonly status: number;

  public constructor(message: string, status: number) {
    super(message);
    this.name = 'ConfigListRequestError';
    this.status = status;
    Object.setPrototypeOf(this, ConfigListRequestError.prototype);
  }
}

export interface IEnsureListResult {
  created: boolean;
  fieldsAdded: string[];
}

const jsonHeaders: { [key: string]: string } = {
  Accept: 'application/json;odata=nometadata',
  'Content-Type': 'application/json;odata=nometadata'
};

const fieldHeaders: { [key: string]: string } = {
  Accept: 'application/json;odata.metadata=none',
  'Content-Type': 'application/json;charset=utf-8',
  'odata-version': ''
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
  const present = new Set(readInternalNames(fieldPayload).map((name) => name.toLowerCase()));
  const fieldsAdded: string[] = [];
  for (const field of CONFIG_FIELD_DEFINITIONS) {
    if (present.has(field.internalName.toLowerCase())) {
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
      const text = await addResponse.text();
      if (isExistingColumnError(text)) {
        continue;
      }
      throw new ConfigListRequestError(sharePointError(addResponse.status, text, listTitle), addResponse.status);
    }
    fieldsAdded.push(field.internalName);
  }
  return { created, fieldsAdded };
}

export async function showConfigList(
  spHttpClient: SPHttpClient,
  siteAbsoluteUrl: string,
  listTitle: string
): Promise<void> {
  const url = buildListUrl(siteAbsoluteUrl, listTitle);
  const headers = {
    ...jsonHeaders,
    'IF-MATCH': '*',
    'X-HTTP-Method': 'MERGE',
    'odata-version': ''
  };
  const visible = await spHttpClient.post(url, SPHttpClient.configurations.v1, {
    headers,
    body: JSON.stringify({ Hidden: false })
  });
  if (!visible.ok && visible.status !== 404) {
    await visible.text();
  }
  const navigation = await spHttpClient.post(url, SPHttpClient.configurations.v1, {
    headers,
    body: JSON.stringify({ OnQuickLaunch: true })
  });
  if (!navigation.ok && navigation.status !== 404) {
    await navigation.text();
  }
}

export async function saveConfigRow(
  spHttpClient: SPHttpClient,
  siteAbsoluteUrl: string,
  listTitle: string,
  draft: IConfigDraft,
  writeCount: boolean
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
        'X-HTTP-Method': 'MERGE',
        'odata-version': ''
      },
    body: JSON.stringify(creating ? buildCreateBody(draft) : buildUpdateBody(draft, writeCount))
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
    return;
  }
  const web = deriveWebFromListUrl(listUrl).webAbsoluteUrl.replace(/\/+$/, '');
  const guid = listGuid.replace(/[{}]/g, '');
  if (await defaultViewHasField(spHttpClient, web, guid, name)) {
    return;
  }
  await addFieldToDefaultView(spHttpClient, web, guid, name);
}

export async function loadFormulaColumns(
  spHttpClient: SPHttpClient,
  webAbsoluteUrl: string,
  listGuid: string
): Promise<IFormulaColumn[]> {
  try {
    const guid = listGuid.replace(/[{}]/g, '');
    const response = await spHttpClient.get(
      `${webAbsoluteUrl.replace(/\/+$/, '')}/_api/web/lists(guid'${guid}')/fields?$select=InternalName,Title,TypeAsString,Hidden&$top=300`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: 'application/json;odata=nometadata' } }
    );
    if (!response.ok) {
      return [];
    }
    const payload = await response.json() as {
      value?: Array<{ InternalName?: string; Title?: string; TypeAsString?: string; Hidden?: boolean }>;
    };
    return formulaColumns(payload.value || []);
  } catch {
    return [];
  }
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

async function defaultViewHasField(
  spHttpClient: SPHttpClient,
  webAbsoluteUrl: string,
  listGuid: string,
  internalName: string
): Promise<boolean> {
  const response = await spHttpClient.get(
    `${webAbsoluteUrl}/_api/web/lists(guid'${listGuid}')/DefaultView/ViewFields`,
    SPHttpClient.configurations.v1,
    { headers: { Accept: 'application/json;odata=nometadata' } }
  );
  if (!response.ok) {
    return false;
  }
  const payload = await response.json() as { Items?: string[] | { results?: string[] } };
  const items = !payload.Items ? [] : (Array.isArray(payload.Items) ? payload.Items : payload.Items.results || []);
  const wanted = internalName.toLowerCase();
  return items.some((item) => item.toLowerCase() === wanted);
}

async function addFieldToDefaultView(
  spHttpClient: SPHttpClient,
  webAbsoluteUrl: string,
  listGuid: string,
  internalName: string
): Promise<void> {
  const response = await spHttpClient.post(
    `${webAbsoluteUrl}/_api/web/lists(guid'${listGuid}')/DefaultView/ViewFields/AddViewField('${internalName}')`,
    SPHttpClient.configurations.v1,
    { headers: { Accept: 'application/json;odata=nometadata' } }
  );
  if (response.ok) {
    return;
  }
  const text = await response.text();
  if (isExistingColumnError(text)) {
    return;
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
