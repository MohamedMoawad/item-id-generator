import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
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
import {
  columnValuesFromCreateBody,
  embedCreateField,
  IPreparedCreate,
  newItemJson,
  readCreateField,
  rewriteNewItemBody
} from './listCreateWatch';

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

interface INumberLock {
  busy: boolean;
  pending: boolean;
}

interface INumberWindow extends Window {
  __autogenNumberLock?: INumberLock;
}

function numberLock(): INumberLock {
  const host = window as INumberWindow;
  if (!host.__autogenNumberLock) {
    host.__autogenNumberLock = { busy: false, pending: false };
  }
  return host.__autogenNumberLock;
}

const CREATE_LOCK_WAIT_MS = 8000;

export async function embedNumberInCreate(
  spHttpClient: SPHttpClient,
  siteAbsoluteUrl: string,
  configListTitle: string,
  webAbsoluteUrl: string,
  listGuid: string,
  listServerRelativeUrl: string,
  body: string
): Promise<IPreparedCreate | undefined> {
  const guid = normalizeGuid(listGuid);
  if (!guid) {
    return undefined;
  }
  const json = newItemJson(body, guid, listServerRelativeUrl);
  if (!json) {
    return undefined;
  }
  const acquired = await acquireNumberLock();
  if (!acquired) {
    return undefined;
  }
  try {
    return await reserveEmbeddedNumber(
      spHttpClient,
      siteAbsoluteUrl,
      configListTitle,
      webAbsoluteUrl,
      guid,
      listServerRelativeUrl,
      body,
      json
    );
  } catch {
    return undefined;
  } finally {
    numberLock().busy = false;
  }
}

async function reserveEmbeddedNumber(
  spHttpClient: SPHttpClient,
  siteAbsoluteUrl: string,
  configListTitle: string,
  webAbsoluteUrl: string,
  listGuid: string,
  listServerRelativeUrl: string,
  body: string,
  json: string
): Promise<IPreparedCreate | undefined> {
  const rows = await loadRows(spHttpClient, siteAbsoluteUrl, configListTitle);
  const config = findConfigForList(rows, listGuid);
  if (!config || !config.isActive || config.id === undefined) {
    return undefined;
  }
  const fieldName = config.numberColumnInternalName.trim();
  if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(fieldName)) {
    return undefined;
  }
  const existing = readCreateField(json, fieldName);
  if (!existing || (existing.found && !isBlankNumber(existing.value))) {
    return undefined;
  }
  const columns = await loadFormulaColumns(spHttpClient, webAbsoluteUrl, listGuid);
  const columnNames = columnNamesInFormula(config.formula);
  const mapped = columnValuesFromCreateBody(json, columnNames, columns);
  if (!mapped || !mapped.complete) {
    return undefined;
  }
  const configItemId = config.id;
  for (let attempt = 1; attempt <= 8; attempt += 1) {
    const embedded = await reserveOneCreate(
      spHttpClient,
      siteAbsoluteUrl,
      configListTitle,
      configItemId,
      fieldName,
      mapped.values,
      body,
      listGuid,
      listServerRelativeUrl
    );
    if (embedded !== 'conflict') {
      return embedded;
    }
  }
  return undefined;
}

async function reserveOneCreate(
  spHttpClient: SPHttpClient,
  siteAbsoluteUrl: string,
  configListTitle: string,
  configItemId: number,
  fieldName: string,
  columns: { [token: string]: string },
  body: string,
  listGuid: string,
  listServerRelativeUrl: string
): Promise<IPreparedCreate | undefined | 'conflict'> {
  const current = await readCounter(spHttpClient, siteAbsoluteUrl, configListTitle, configItemId);
  if (!current || !isActiveValue(current.IsActive) || !current.Formula) {
    return undefined;
  }
  const previousCount = Number(current.CurrentCount) || 0;
  const previousReset = current.LastResetDate || '';
  const padLength = Number(current.PadLength);
  let planned: { code: string; sequence: number; lastResetDate: string };
  try {
    planned = planNextCode({
      formula: current.Formula,
      resetPeriod: current.ResetPeriod || 'None',
      lastResetDate: current.LastResetDate,
      currentCount: previousCount,
      padLength: Number.isInteger(padLength) && padLength >= 0 ? padLength : 0,
      now: new Date(),
      columns
    });
  } catch {
    return undefined;
  }
  const swapped = await swapCounter(
    spHttpClient,
    siteAbsoluteUrl,
    configListTitle,
    configItemId,
    etagOf(current),
    planned.sequence,
    planned.lastResetDate
  );
  if (swapped !== 'swapped') {
    return swapped === 'conflict' ? 'conflict' : undefined;
  }
  const rewritten = rewriteNewItemBody(body, listGuid, listServerRelativeUrl, (part) => embedCreateField(part, fieldName, planned.code));
  if (!rewritten) {
    await releaseReservedCount(
      spHttpClient,
      siteAbsoluteUrl,
      configListTitle,
      configItemId,
      planned.sequence,
      previousCount,
      previousReset
    );
    return undefined;
  }
  return {
    body: rewritten,
    rollback: () => releaseReservedCount(
      spHttpClient,
      siteAbsoluteUrl,
      configListTitle,
      configItemId,
      planned.sequence,
      previousCount,
      previousReset
    )
  };
}

async function acquireNumberLock(): Promise<boolean> {
  const started = Date.now();
  let lock = numberLock();
  while (lock.busy) {
    if (Date.now() - started >= CREATE_LOCK_WAIT_MS) {
      return false;
    }
    await delay(40);
    lock = numberLock();
  }
  lock.busy = true;
  return true;
}

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => {
    window.setTimeout(resolve, ms);
  });
}

export async function numberBlankItems(
  spHttpClient: SPHttpClient,
  siteAbsoluteUrl: string,
  configListTitle: string,
  webAbsoluteUrl: string,
  listGuid: string
): Promise<IAssignedNumber[]> {
  const lock = numberLock();
  if (lock.busy) {
    lock.pending = true;
    return [];
  }
  lock.busy = true;
  let assigned: IAssignedNumber[] = [];
  try {
    assigned = await numberBlankItemsNow(spHttpClient, siteAbsoluteUrl, configListTitle, webAbsoluteUrl, listGuid);
  } finally {
    lock.busy = false;
  }
  if (!lock.pending) {
    return assigned;
  }
  lock.pending = false;
  const more = await numberBlankItems(spHttpClient, siteAbsoluteUrl, configListTitle, webAbsoluteUrl, listGuid);
  return assigned.concat(more);
}

async function numberBlankItemsNow(
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
    if (code === '') {
      continue;
    }
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
    return '';
  }
  const columnValues = columnMapForItem(item, columnNames, columns);
  for (let attempt = 1; attempt <= 8; attempt += 1) {
    const fresh = await readItem(spHttpClient, webAbsoluteUrl, listGuid, itemId, ['Id', fieldName], []);
    if (!fresh || !isBlankNumber(fresh[fieldName])) {
      return '';
    }
    const current = await readCounter(spHttpClient, siteAbsoluteUrl, configListTitle, configItemId);
    if (!current || !isActiveValue(current.IsActive) || !current.Formula) {
      return undefined;
    }
    const previousCount = Number(current.CurrentCount) || 0;
    const previousReset = current.LastResetDate || '';
    const padLength = Number(current.PadLength);
    const planned = planNextCode({
      formula: current.Formula,
      resetPeriod: current.ResetPeriod || 'None',
      lastResetDate: current.LastResetDate,
      currentCount: previousCount,
      padLength: Number.isInteger(padLength) && padLength >= 0 ? padLength : 0,
      now: new Date(),
      columns: columnValues
    });
    const swapped = await swapCounter(
      spHttpClient,
      siteAbsoluteUrl,
      configListTitle,
      configItemId,
      etagOf(current),
      planned.sequence,
      planned.lastResetDate
    );
    if (swapped === 'conflict') {
      continue;
    }
    if (swapped === 'failed') {
      return undefined;
    }
    const wrote = await writeNumber(
      spHttpClient,
      webAbsoluteUrl,
      listGuid,
      itemId,
      fieldName,
      planned.code,
      etagOf(fresh)
    );
    if (wrote === 'wrote') {
      return planned.code;
    }
    if (wrote === 'conflict') {
      const again = await readItem(spHttpClient, webAbsoluteUrl, listGuid, itemId, ['Id', fieldName], []);
      if (again && !isBlankNumber(again[fieldName])) {
        await releaseReservedCount(
          spHttpClient,
          siteAbsoluteUrl,
          configListTitle,
          configItemId,
          planned.sequence,
          previousCount,
          previousReset
        );
        return '';
      }
      const retry = await writeNumber(
        spHttpClient,
        webAbsoluteUrl,
        listGuid,
        itemId,
        fieldName,
        planned.code,
        again ? etagOf(again) : '*'
      );
      return retry === 'wrote' ? planned.code : undefined;
    }
    return undefined;
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
  const payload = await response.json() as IListItemRow;
  rememberEtag(response, payload);
  return payload;
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
  const payload = await response.json() as ICounterRow;
  rememberEtag(response, payload);
  return payload;
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
  code: string,
  etag: string
): Promise<'wrote' | 'conflict' | 'failed'> {
  const body: { [key: string]: string } = {};
  body[fieldName] = code;
  const response = await spHttpClient.post(
    `${trimSlash(webAbsoluteUrl)}/_api/web/lists(guid'${listGuid}')/items(${itemId})`,
    SPHttpClient.configurations.v1,
    {
      headers: {
        ...mergeHeaders,
        'IF-MATCH': etag || '*'
      },
      body: JSON.stringify(body)
    }
  );
  if (response.ok || response.status === 204) {
    return 'wrote';
  }
  if (response.status === 412 || response.status === 409) {
    return 'conflict';
  }
  return 'failed';
}

async function releaseReservedCount(
  spHttpClient: SPHttpClient,
  siteAbsoluteUrl: string,
  configListTitle: string,
  itemId: number,
  reservedSequence: number,
  previousCount: number,
  previousReset: string
): Promise<void> {
  const current = await readCounter(spHttpClient, siteAbsoluteUrl, configListTitle, itemId);
  if (!current || Number(current.CurrentCount) !== reservedSequence) {
    return;
  }
  await swapCounter(
    spHttpClient,
    siteAbsoluteUrl,
    configListTitle,
    itemId,
    etagOf(current),
    previousCount,
    previousReset
  );
}

function etagOf(row: IListItemRow | ICounterRow): string {
  const record = row as { 'odata.etag'?: string; '@odata.etag'?: string };
  return record['odata.etag'] || record['@odata.etag'] || '*';
}

function rememberEtag(response: SPHttpClientResponse, row: IListItemRow | ICounterRow): void {
  if (etagOf(row) !== '*') {
    return;
  }
  const header = response.headers.get('ETag') || response.headers.get('etag') || '';
  if (header) {
    (row as IListItemRow)['odata.etag'] = header;
  }
}

function configItems(siteAbsoluteUrl: string, configListTitle: string): string {
  return `${trimSlash(siteAbsoluteUrl)}/_api/web/lists/getbytitle('${escapeListTitle(configListTitle)}')/items`;
}

function trimSlash(value: string): string {
  return value.replace(/\/+$/, '');
}
