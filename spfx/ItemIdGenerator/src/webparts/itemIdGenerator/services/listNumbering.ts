import { SPHttpClient } from '@microsoft/sp-http';
import {
  escapeListTitle,
  findConfigForList,
  isActiveValue,
  normalizeGuid
} from './configContract';
import { loadConfigRows, loadFormulaColumns } from './configListClient';
import {
  columnMapForItem,
  columnNamesInFormula,
  IFormulaColumn,
  isBlankNumber,
  itemQueryParts,
  planNextCode
} from './formula';

const metadataHeaders: { [key: string]: string } = {
  Accept: 'application/json;odata=minimalmetadata'
};

const mergeHeaders: { [key: string]: string } = {
  Accept: 'application/json;odata=nometadata',
  'Content-Type': 'application/json;odata=nometadata',
  'IF-MATCH': '*',
  'X-HTTP-Method': 'MERGE'
};

interface ICounterRow {
  Id?: number;
  Formula?: string;
  CurrentCount?: number;
  ResetPeriod?: string;
  LastResetDate?: string;
  IsActive?: boolean | string | number;
  PadLength?: number | string;
  NumberColumnInternalName?: string;
  'odata.etag'?: string;
}

interface IListItemRow {
  Id?: number;
  [key: string]: unknown;
}

export interface IAssignedNumber {
  itemId: number;
  fieldName: string;
  code: string;
}

export async function numberBlankItems(
  spHttpClient: SPHttpClient,
  siteAbsoluteUrl: string,
  configListTitle: string,
  webAbsoluteUrl: string,
  listGuid: string
): Promise<IAssignedNumber[]> {
  const guid = normalizeGuid(listGuid);
  if (!guid) {
    return [];
  }
  const rows = await loadRows(spHttpClient, siteAbsoluteUrl, configListTitle);
  const config = findConfigForList(rows, guid);
  if (!config || !config.isActive || config.id === undefined) {
    return [];
  }
  const fieldName = config.numberColumnInternalName.trim();
  if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(fieldName)) {
    return [];
  }
  const columns = await loadFormulaColumns(spHttpClient, webAbsoluteUrl, guid);
  const columnNames = columnNamesInFormula(config.formula);
  const items = await readRecentItems(spHttpClient, webAbsoluteUrl, guid, fieldName);
  const blanks = items.filter((item) => item.Id !== undefined && isBlankNumber(item[fieldName])).slice(0, 20);
  const assigned: IAssignedNumber[] = [];
  for (const item of blanks) {
    const code = await reserveAndWrite(
      spHttpClient,
      siteAbsoluteUrl,
      configListTitle,
      webAbsoluteUrl,
      guid,
      config.id,
      fieldName,
      item.Id as number,
      columnNames,
      columns
    );
    if (!code) {
      break;
    }
    assigned.push({ itemId: item.Id as number, fieldName, code });
  }
  return assigned;
}

async function reserveAndWrite(
  spHttpClient: SPHttpClient,
  siteAbsoluteUrl: string,
  configListTitle: string,
  webAbsoluteUrl: string,
  listGuid: string,
  configItemId: number,
  fieldName: string,
  itemId: number,
  columnNames: string[],
  columns: IFormulaColumn[]
): Promise<string | undefined> {
  const item = await readItemForFormula(spHttpClient, webAbsoluteUrl, listGuid, itemId, fieldName, columnNames, columns);
  if (!item || !isBlankNumber(item[fieldName])) {
    return undefined;
  }
  const columnValues = columnMapForItem(item, columnNames, columns);
  for (let attempt = 1; attempt <= 8; attempt += 1) {
    const current = await readCounter(spHttpClient, siteAbsoluteUrl, configListTitle, configItemId);
    if (!current || !isActiveValue(current.IsActive) || !current.Formula) {
      return undefined;
    }
    const padLength = Number(current.PadLength);
    const planned = planNextCode({
      formula: current.Formula,
      resetPeriod: current.ResetPeriod || 'None',
      lastResetDate: current.LastResetDate,
      currentCount: Number(current.CurrentCount) || 0,
      padLength: Number.isInteger(padLength) && padLength >= 0 ? padLength : 0,
      now: new Date(),
      columns: columnValues
    });
    const swapped = await swapCounter(
      spHttpClient,
      siteAbsoluteUrl,
      configListTitle,
      configItemId,
      current['odata.etag'] || '*',
      planned.sequence,
      planned.lastResetDate
    );
    if (swapped === 'conflict') {
      continue;
    }
    if (swapped === 'failed') {
      return undefined;
    }
    const wrote = await writeNumber(spHttpClient, webAbsoluteUrl, listGuid, itemId, fieldName, planned.code);
    return wrote ? planned.code : undefined;
  }
  return undefined;
}

async function loadRows(
  spHttpClient: SPHttpClient,
  siteAbsoluteUrl: string,
  configListTitle: string
): Promise<Awaited<ReturnType<typeof loadConfigRows>>> {
  return loadConfigRows(spHttpClient, siteAbsoluteUrl, configListTitle);
}

async function readRecentItems(
  spHttpClient: SPHttpClient,
  webAbsoluteUrl: string,
  listGuid: string,
  fieldName: string
): Promise<IListItemRow[]> {
  const url = `${trimSlash(webAbsoluteUrl)}/_api/web/lists(guid'${listGuid}')/items?$select=Id,${fieldName}&$orderby=Id desc&$top=30`;
  const response = await spHttpClient.get(url, SPHttpClient.configurations.v1, { headers: metadataHeaders });
  if (!response.ok) {
    return [];
  }
  const payload = await response.json() as { value?: IListItemRow[] };
  return payload.value || [];
}

async function readItemForFormula(
  spHttpClient: SPHttpClient,
  webAbsoluteUrl: string,
  listGuid: string,
  itemId: number,
  fieldName: string,
  columnNames: string[],
  columns: IFormulaColumn[]
): Promise<IListItemRow | undefined> {
  const parts = itemQueryParts(fieldName, columnNames, columns);
  const rich = await readItem(spHttpClient, webAbsoluteUrl, listGuid, itemId, parts.select, parts.expand);
  if (rich) {
    return rich;
  }
  return readItem(spHttpClient, webAbsoluteUrl, listGuid, itemId, ['Id', fieldName], []);
}

async function readItem(
  spHttpClient: SPHttpClient,
  webAbsoluteUrl: string,
  listGuid: string,
  itemId: number,
  select: string[],
  expand: string[]
): Promise<IListItemRow | undefined> {
  const expandQuery = expand.length > 0 ? `&$expand=${expand.join(',')}` : '';
  const url = `${trimSlash(webAbsoluteUrl)}/_api/web/lists(guid'${listGuid}')/items(${itemId})?$select=${select.join(',')}${expandQuery}`;
  const response = await spHttpClient.get(url, SPHttpClient.configurations.v1, { headers: metadataHeaders });
  if (!response.ok) {
    return undefined;
  }
  return await response.json() as IListItemRow;
}

async function readCounter(
  spHttpClient: SPHttpClient,
  siteAbsoluteUrl: string,
  configListTitle: string,
  itemId: number
): Promise<ICounterRow | undefined> {
  const url = `${configItems(siteAbsoluteUrl, configListTitle)}(${itemId})?$select=Id,Formula,CurrentCount,ResetPeriod,LastResetDate,IsActive,PadLength,NumberColumnInternalName`;
  const response = await spHttpClient.get(url, SPHttpClient.configurations.v1, { headers: metadataHeaders });
  if (!response.ok) {
    return undefined;
  }
  return await response.json() as ICounterRow;
}

async function swapCounter(
  spHttpClient: SPHttpClient,
  siteAbsoluteUrl: string,
  configListTitle: string,
  itemId: number,
  etag: string,
  sequence: number,
  lastResetDate: string
): Promise<'swapped' | 'conflict' | 'failed'> {
  const response = await spHttpClient.post(
    `${configItems(siteAbsoluteUrl, configListTitle)}(${itemId})`,
    SPHttpClient.configurations.v1,
    {
      headers: {
        ...mergeHeaders,
        'IF-MATCH': etag
      },
      body: JSON.stringify({
        CurrentCount: sequence,
        LastResetDate: lastResetDate
      })
    }
  );
  if (response.ok || response.status === 204) {
    return 'swapped';
  }
  if (response.status === 412 || response.status === 409) {
    return 'conflict';
  }
  return 'failed';
}

async function writeNumber(
  spHttpClient: SPHttpClient,
  webAbsoluteUrl: string,
  listGuid: string,
  itemId: number,
  fieldName: string,
  code: string
): Promise<boolean> {
  const body: { [key: string]: string } = {};
  body[fieldName] = code;
  const response = await spHttpClient.post(
    `${trimSlash(webAbsoluteUrl)}/_api/web/lists(guid'${listGuid}')/items(${itemId})`,
    SPHttpClient.configurations.v1,
    {
      headers: mergeHeaders,
      body: JSON.stringify(body)
    }
  );
  return response.ok || response.status === 204;
}

function configItems(siteAbsoluteUrl: string, configListTitle: string): string {
  return `${trimSlash(siteAbsoluteUrl)}/_api/web/lists/getbytitle('${escapeListTitle(configListTitle)}')/items`;
}

function trimSlash(value: string): string {
  return value.replace(/\/+$/, '');
}
