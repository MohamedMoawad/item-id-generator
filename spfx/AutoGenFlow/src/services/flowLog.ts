import { SPHttpClient } from '@microsoft/sp-http';
import {
  buildCreateFieldUrl,
  buildCreateItemUrl,
  buildCreateListUrl,
  buildFieldsUrl,
  buildListUrl,
  DEFAULT_LOG_LIST_TITLE,
  isExistingColumnError
} from './contract';
import { buildCreateFieldBody, IConfigFieldDefinition } from './fields';

export interface IFlowLogEntry {
  step: string;
  status: 'Completed' | 'Failed';
  detail: string;
  listTitle: string;
  listGuid: string;
  itemId?: number;
  code?: string;
}

const LOG_FIELDS: IConfigFieldDefinition[] = [
  {
    internalName: 'LogStatus',
    displayName: 'Status',
    schemaXml: '<Field Type="Text" Name="LogStatus" StaticName="LogStatus" DisplayName="Status" />'
  },
  {
    internalName: 'Detail',
    displayName: 'Detail',
    schemaXml: '<Field Type="Note" Name="Detail" StaticName="Detail" DisplayName="Detail" NumLines="6" RichText="FALSE" />'
  },
  {
    internalName: 'TargetListTitle',
    displayName: 'Target list title',
    schemaXml: '<Field Type="Text" Name="TargetListTitle" StaticName="TargetListTitle" DisplayName="Target list title" />'
  },
  {
    internalName: 'TargetListGuid',
    displayName: 'Target list GUID',
    schemaXml: '<Field Type="Text" Name="TargetListGuid" StaticName="TargetListGuid" DisplayName="Target list GUID" />'
  },
  {
    internalName: 'SourceItemId',
    displayName: 'Item id',
    schemaXml: '<Field Type="Number" Name="SourceItemId" StaticName="SourceItemId" DisplayName="Item id" Decimals="0" Min="0" />'
  },
  {
    internalName: 'NumberCode',
    displayName: 'Number',
    schemaXml: '<Field Type="Text" Name="NumberCode" StaticName="NumberCode" DisplayName="Number" />'
  }
];

const jsonHeaders: { [name: string]: string } = {
  Accept: 'application/json;odata=nometadata',
  'Content-Type': 'application/json;odata=nometadata'
};

const fieldHeaders: { [name: string]: string } = {
  Accept: 'application/json;odata.metadata=none',
  'Content-Type': 'application/json;charset=utf-8',
  'odata-version': ''
};

export async function ensureLogList(spHttpClient: SPHttpClient, siteAbsoluteUrl: string): Promise<void> {
  const title = DEFAULT_LOG_LIST_TITLE;
  const existing = await spHttpClient.get(
    buildListUrl(siteAbsoluteUrl, title),
    SPHttpClient.configurations.v1,
    { headers: { Accept: 'application/json;odata=nometadata' } }
  );
  if (existing.status === 404) {
    const created = await spHttpClient.post(buildCreateListUrl(siteAbsoluteUrl), SPHttpClient.configurations.v1, {
      headers: jsonHeaders,
      body: JSON.stringify({
        BaseTemplate: 100,
        Title: title,
        Description: 'AutoGen Flow writes one row for each completed step. The newest row is the last step that ran.',
        Hidden: false
      })
    });
    if (!created.ok) {
      await created.text();
      return;
    }
  } else if (!existing.ok) {
    await existing.text();
    return;
  } else {
    await existing.text();
  }
  await addMissingFields(spHttpClient, siteAbsoluteUrl, title);
}

export async function showLogList(spHttpClient: SPHttpClient, siteAbsoluteUrl: string): Promise<void> {
  const url = buildListUrl(siteAbsoluteUrl, DEFAULT_LOG_LIST_TITLE);
  const headers = {
    ...jsonHeaders,
    'IF-MATCH': '*',
    'X-HTTP-Method': 'MERGE',
    'odata-version': ''
  };
  const visible = await spHttpClient.post(url, SPHttpClient.configurations.v1, {
    headers,
    body: JSON.stringify({ Hidden: false, OnQuickLaunch: true })
  });
  if (!visible.ok && visible.status !== 404) {
    await visible.text();
  }
}

export async function writeFlowLog(
  spHttpClient: SPHttpClient,
  siteAbsoluteUrl: string,
  entry: IFlowLogEntry
): Promise<void> {
  try {
    await ensureLogList(spHttpClient, siteAbsoluteUrl);
    const body: { [name: string]: string | number } = {
      Title: clip(entry.step, 255),
      LogStatus: entry.status,
      Detail: clip(entry.detail, 2000),
      TargetListTitle: clip(entry.listTitle, 255),
      TargetListGuid: clip(entry.listGuid, 255),
      NumberCode: clip(entry.code || '', 255)
    };
    if (entry.itemId !== undefined && entry.itemId > 0) {
      body.SourceItemId = entry.itemId;
    }
    const response = await spHttpClient.post(
      buildCreateItemUrl(siteAbsoluteUrl, DEFAULT_LOG_LIST_TITLE),
      SPHttpClient.configurations.v1,
      { headers: jsonHeaders, body: JSON.stringify(body) }
    );
    await response.text();
  } catch {
    return;
  }
}

async function addMissingFields(spHttpClient: SPHttpClient, siteAbsoluteUrl: string, title: string): Promise<void> {
  const fieldsResponse = await spHttpClient.get(
    buildFieldsUrl(siteAbsoluteUrl, title),
    SPHttpClient.configurations.v1,
    { headers: { Accept: 'application/json;odata=nometadata' } }
  );
  const fieldsText = await fieldsResponse.text();
  if (!fieldsResponse.ok) {
    return;
  }
  const payload = JSON.parse(fieldsText) as {
    value?: { InternalName?: string }[];
    d?: { results?: { InternalName?: string }[] };
  };
  const names = (payload.value || (payload.d && payload.d.results) || [])
    .map((field) => (field.InternalName || '').toLowerCase());
  for (let index = 0; index < LOG_FIELDS.length; index += 1) {
    const field = LOG_FIELDS[index];
    if (names.indexOf(field.internalName.toLowerCase()) >= 0) {
      continue;
    }
    const addResponse = await spHttpClient.post(
      buildCreateFieldUrl(siteAbsoluteUrl, title),
      SPHttpClient.configurations.v1,
      { headers: fieldHeaders, body: JSON.stringify(buildCreateFieldBody(field.schemaXml)) }
    );
    const addText = await addResponse.text();
    if (!addResponse.ok && !isExistingColumnError(addText)) {
      return;
    }
  }
}

function clip(value: string, length: number): string {
  const text = value || '';
  return text.length <= length ? text : text.slice(0, length);
}
