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
import { counterIsCurrent, planContinuousNumbers } from './sequencePlan';
import {
  columnValuesFromCreateBody,
  createSaveFailed,
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
  'X-HTTP-Method': 'MERGE',
  'odata-version': ''
};

const validateHeaders: { [key: string]: string } = {
  Accept: 'application/json;odata=nometadata',
  'Content-Type': 'application/json;odata=nometadata',
  'odata-version': ''
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
  followUp: boolean;
  quietUntil: number;
}

interface INumberWindow extends Window {
  __autogenNumberLock?: INumberLock;
}

function numberLock(): INumberLock {
  const host = window as INumberWindow;
  if (!host.__autogenNumberLock) {
    host.__autogenNumberLock = { busy: false, pending: false, followUp: false, quietUntil: 0 };
  }
  return host.__autogenNumberLock;
}

const CREATE_LOCK_WAIT_MS = 60000;
const ITEM_PAGE_SIZE = 200;
const ITEM_PAGE_LIMIT = 50;
const WRITE_CHUNK = 8;
const QUIET_MS = 10000;

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
    releaseNumberLock(spHttpClient, siteAbsoluteUrl, configListTitle, webAbsoluteUrl, guid);
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

function releaseNumberLock(
  spHttpClient: SPHttpClient,
  siteAbsoluteUrl: string,
  configListTitle: string,
  webAbsoluteUrl: string,
  listGuid: string
): void {
  const lock = numberLock();
  lock.busy = false;
  if (!lock.pending) {
    return;
  }
  lock.pending = false;
  numberBlankItems(spHttpClient, siteAbsoluteUrl, configListTitle, webAbsoluteUrl, listGuid, true).catch(() => undefined);
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
  listGuid: string,
  force = false,
  pass = 0
): Promise<IAssignedNumber[]> {
  const lock = numberLock();
  if (lock.busy) {
    lock.pending = true;
    return [];
  }
  if (!force && !lock.pending && Date.now() < lock.quietUntil) {
    return [];
  }
  lock.busy = true;
  lock.pending = false;
  let assigned: IAssignedNumber[] = [];
  try {
    assigned = await numberBlankItemsNow(spHttpClient, siteAbsoluteUrl, configListTitle, webAbsoluteUrl, listGuid);
  } finally {
    lock.busy = false;
  }
  const again = lock.pending || lock.followUp;
  lock.followUp = false;
  if (!again || pass >= 3) {
    return assigned;
  }
  const more = await numberBlankItems(spHttpClient, siteAbsoluteUrl, configListTitle, webAbsoluteUrl, listGuid, true, pass + 1);
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
  const padLength = Number(config.padLength);
  const now = new Date();
  const loaded = await readItemsInOrder(spHttpClient, webAbsoluteUrl, guid, fieldName, columnNames, columns);
  if (!loaded) {
    markFollowUp();
    return [];
  }
  let planned: ReturnType<typeof planContinuousNumbers>;
  try {
    planned = planContinuousNumbers({
      items: loaded.items,
      formula: config.formula,
      resetPeriod: config.resetPeriod,
      padLength: Number.isInteger(padLength) && padLength >= 0 ? padLength : 0,
      now
    });
  } catch {
    markFollowUp();
    return [];
  }
  const changed = planned.plans.filter((item) => item.changed);
  const assigned: IAssignedNumber[] = [];
  let failed = false;
  for (let index = 0; index < changed.length; index += WRITE_CHUNK) {
    const chunk = changed.slice(index, index + WRITE_CHUNK);
    const writes: Array<Promise<'wrote' | 'conflict' | 'failed'>> = [];
    for (let itemIndex = 0; itemIndex < chunk.length; itemIndex += 1) {
      const item = chunk[itemIndex];
      writes.push(writeNumber(spHttpClient, webAbsoluteUrl, guid, item.id, fieldName, item.code, '*'));
    }
    const results = await Promise.all(writes);
    for (let itemIndex = 0; itemIndex < chunk.length; itemIndex += 1) {
      if (results[itemIndex] === 'wrote') {
        assigned.push({ itemId: chunk[itemIndex].id, fieldName, code: chunk[itemIndex].code });
      } else {
        failed = true;
      }
    }
  }
  const counter = await readCounter(spHttpClient, siteAbsoluteUrl, configListTitle, config.id);
  const aligned = !!counter && counterIsCurrent(
    Number(counter.CurrentCount) || 0,
    counter.LastResetDate,
    planned.currentCount,
    config.resetPeriod,
    now
  );
  if (!counter) {
    failed = true;
  } else if (!aligned) {
    const saved = await setCounterTo(
      spHttpClient,
      siteAbsoluteUrl,
      configListTitle,
      config.id,
      planned.currentCount,
      planned.lastResetDate,
      config.resetPeriod,
      now
    );
    if (!saved) {
      failed = true;
    }
  }
  const lock = numberLock();
  lock.followUp = failed || !loaded.complete || assigned.length > 0;
  lock.quietUntil = lock.followUp ? 0 : Date.now() + QUIET_MS;
  return assigned;
}

function markFollowUp(): void {
  const lock = numberLock();
  lock.followUp = true;
  lock.quietUntil = 0;
}

async function loadRows(
  spHttpClient: SPHttpClient,
  siteAbsoluteUrl: string,
  configListTitle: string
): Promise<Awaited<ReturnType<typeof loadConfigRows>>> {
  return loadConfigRows(spHttpClient, siteAbsoluteUrl, configListTitle);
}

interface IOrderedItems {
  items: Array<{ id: number; created: string; current: string; columns: { [token: string]: string } }>;
  complete: boolean;
}

async function readItemsInOrder(
  spHttpClient: SPHttpClient,
  webAbsoluteUrl: string,
  listGuid: string,
  fieldName: string,
  columnNames: string[],
  columns: IFormulaColumn[]
): Promise<IOrderedItems | undefined> {
  const parts = itemQueryParts(fieldName, columnNames, columns);
  if (parts.select.indexOf('Created') < 0) {
    parts.select.push('Created');
  }
  const expandQuery = parts.expand.length > 0 ? `&$expand=${parts.expand.join(',')}` : '';
  const items: IOrderedItems['items'] = [];
  let lastId = 0;
  for (let page = 0; page < ITEM_PAGE_LIMIT; page += 1) {
    const url = `${trimSlash(webAbsoluteUrl)}/_api/web/lists(guid'${listGuid}')/items?$select=${parts.select.join(',')}&$orderby=Id asc&$filter=Id gt ${lastId}&$top=${ITEM_PAGE_SIZE}${expandQuery}`;
    const response = await spHttpClient.get(url, SPHttpClient.configurations.v1, { headers: metadataHeaders });
    if (!response.ok) {
      return page === 0 ? undefined : { items, complete: false };
    }
    const payload = await response.json() as { value?: IListItemRow[] };
    const rows = payload.value || [];
    for (let index = 0; index < rows.length; index += 1) {
      const row = rows[index];
      const id = Number(row.Id);
      if (!Number.isInteger(id) || id <= lastId) {
        continue;
      }
      items.push({
        id,
        created: sharePointDate(row.Created),
        current: cellText(row[fieldName]),
        columns: columnMapForItem(row, columnNames, columns)
      });
      lastId = id;
    }
    if (rows.length < ITEM_PAGE_SIZE) {
      return { items, complete: true };
    }
  }
  return { items, complete: false };
}

async function setCounterTo(
  spHttpClient: SPHttpClient,
  siteAbsoluteUrl: string,
  configListTitle: string,
  itemId: number,
  sequence: number,
  lastResetDate: string,
  resetPeriod: string,
  now: Date
): Promise<boolean> {
  for (let attempt = 1; attempt <= 5; attempt += 1) {
    const current = await readCounter(spHttpClient, siteAbsoluteUrl, configListTitle, itemId);
    if (!current) {
      return false;
    }
    if (counterIsCurrent(Number(current.CurrentCount) || 0, current.LastResetDate, sequence, resetPeriod, now)) {
      return true;
    }
    const swapped = await swapCounter(
      spHttpClient,
      siteAbsoluteUrl,
      configListTitle,
      itemId,
      etagOf(current),
      sequence,
      lastResetDate
    );
    if (swapped === 'swapped') {
      return true;
    }
    if (swapped === 'failed') {
      return false;
    }
  }
  return false;
}

function cellText(value: unknown): string {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).trim();
}

function sharePointDate(value: unknown): string {
  if (typeof value !== 'string') {
    return '';
  }
  const ticks = /^\/Date\((-?\d+)\)\/$/.exec(value);
  if (!ticks) {
    return value;
  }
  const parsed = new Date(Number(ticks[1]));
  return Number.isNaN(parsed.getTime()) ? '' : parsed.toISOString();
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
  const validated = await writeNumberByValidate(spHttpClient, webAbsoluteUrl, listGuid, itemId, fieldName, code);
  if (validated === 'wrote') {
    return 'wrote';
  }
  return writeNumberByMerge(spHttpClient, webAbsoluteUrl, listGuid, itemId, fieldName, code, etag);
}

async function writeNumberByValidate(
  spHttpClient: SPHttpClient,
  webAbsoluteUrl: string,
  listGuid: string,
  itemId: number,
  fieldName: string,
  code: string
): Promise<'wrote' | 'failed'> {
  const response = await spHttpClient.post(
    `${trimSlash(webAbsoluteUrl)}/_api/web/lists(guid'${listGuid}')/items(${itemId})/ValidateUpdateListItem()`,
    SPHttpClient.configurations.v1,
    {
      headers: validateHeaders,
      body: JSON.stringify({
        formValues: [{ FieldName: fieldName, FieldValue: code }],
        bNewDocumentUpdate: false
      })
    }
  );
  if (!response.ok) {
    return 'failed';
  }
  const payload: unknown = await response.json().catch(() => undefined);
  return createSaveFailed(payload) ? 'failed' : 'wrote';
}

async function writeNumberByMerge(
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
