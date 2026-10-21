import { SPHttpClient } from '@microsoft/sp-http';
import { findConfigForList, IConfigDraft, normalizeGuid, wholeCount } from './contract';
import { loadConfigRows } from './configListClient';
import { buildItemQuery } from './itemQuery';
import { planNextCode } from './formula';
import { isBlankNumber, oldestFirst, readColumnValues } from './listFillValues';
import { writeFlowLog } from './flowLog';

export { isBlankNumber, readColumnValues };

export interface IListFillTarget {
  spHttpClient: SPHttpClient;
  siteAbsoluteUrl: string;
  webAbsoluteUrl: string;
  configListTitle: string;
  listGuid: string;
  listTitle?: string;
}

interface IFillWindow extends Window {
  __autogenFlowFill?: number;
  __autogenFlowFillTarget?: IListFillTarget;
}

const jsonHeaders: { [name: string]: string } = {
  Accept: 'application/json;odata=nometadata',
  'Content-Type': 'application/json;odata=nometadata',
  'IF-MATCH': '*',
  'X-HTTP-Method': 'MERGE',
  'odata-version': ''
};

let running = false;

export function watchListNumbers(target: IListFillTarget | undefined): void {
  const host = window as IFillWindow;
  host.__autogenFlowFillTarget = target;
  if (!target) {
    return;
  }
  fillBlankNumbers(target).catch(() => undefined);
  if (host.__autogenFlowFill !== undefined) {
    return;
  }
  host.__autogenFlowFill = window.setInterval(() => {
    const current = host.__autogenFlowFillTarget;
    if (!current) {
      return;
    }
    fillBlankNumbers(current).catch(() => undefined);
  }, 2000);
}

export async function fillBlankNumbers(target: IListFillTarget): Promise<number> {
  if (running) {
    return 0;
  }
  running = true;
  const pass = `Pass ${new Date().toISOString()}`;
  try {
    const rows = await loadConfigRows(target.spHttpClient, target.siteAbsoluteUrl, target.configListTitle);
    const config = findConfigForList(rows, target.listGuid);
    if (!config || !config.isActive || config.id === undefined) {
      return 0;
    }
    const field = config.numberColumnInternalName.trim();
    if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(field)) {
      await trace(target, 'Load rule', 'Failed', `${pass}. The number column name cannot be written.`, undefined, '');
      return 0;
    }
    const loaded = await loadItems(target, config, field);
    if (loaded.error) {
      await trace(target, 'Load items', 'Failed', `${pass}. ${loaded.error}`, undefined, '');
      return 0;
    }
    const blanks = loaded.items.filter((item) => isBlankNumber(item[field]));
    if (blanks.length === 0) {
      return 0;
    }
    await trace(target, 'Load rule', 'Completed', `${pass}. Column ${field}. Formula ${config.formula}. Count ${wholeCount(config.currentCount)}.`, undefined, '');
    await trace(target, 'Load items', 'Completed', `${pass}. ${loaded.items.length} recent items. ${blanks.length} blank.`, undefined, '');
    let count = wholeCount(config.currentCount);
    let lastReset = config.lastResetDate;
    let wrote = 0;
    for (let index = 0; index < loaded.items.length; index += 1) {
      const item = loaded.items[index];
      if (!isBlankNumber(item[field])) {
        continue;
      }
      const planned = planNextCode({
        formula: config.formula,
        resetPeriod: config.resetPeriod,
        lastResetDate: lastReset,
        currentCount: count,
        padLength: config.padLength,
        now: createdDate(item.Created),
        columns: readColumnValues(item, config.formulaColumns.split(',').filter((token) => token.trim().length > 0))
      });
      await trace(target, 'Plan number', 'Completed', `${pass}. Item ${item.Id}.`, item.Id, planned.code);
      const itemSaved = await mergeItem(target.spHttpClient, itemUrl(target, item.Id), { [field]: planned.code });
      if (!itemSaved.ok) {
        await trace(target, 'Write number', 'Failed', `${pass}. Item ${item.Id}. ${itemSaved.detail}`, item.Id, planned.code);
        break;
      }
      await trace(target, 'Write number', 'Completed', `${pass}. Item ${item.Id}.`, item.Id, planned.code);
      const countSaved = await mergeItem(
        target.spHttpClient,
        configUrl(target, config.id),
        { CurrentCount: planned.sequence, LastResetDate: planned.lastResetDate }
      );
      if (!countSaved.ok) {
        await trace(target, 'Save count', 'Failed', `${pass}. Item ${item.Id} has ${planned.code}. The stored count stayed ${count}. ${countSaved.detail}`, item.Id, planned.code);
        break;
      }
      await trace(target, 'Save count', 'Completed', `${pass}. Count is now ${planned.sequence}.`, item.Id, planned.code);
      count = planned.sequence;
      lastReset = planned.lastResetDate;
      wrote += 1;
    }
    if (wrote > 0) {
      const refreshed = clickRefresh();
      await trace(target, 'Refresh list', refreshed ? 'Completed' : 'Failed', `${pass}. Wrote ${wrote}.`, undefined, '');
    }
    return wrote;
  } catch (error) {
    const message = error instanceof Error ? error.message : 'The numbering pass stopped.';
    await trace(target, 'Numbering pass', 'Failed', `${pass}. ${message}`, undefined, '');
    return 0;
  } finally {
    releaseFill();
  }
}

async function trace(
  target: IListFillTarget,
  step: string,
  status: 'Completed' | 'Failed',
  detail: string,
  itemId: number | undefined,
  code: string
): Promise<void> {
  await writeFlowLog(target.spHttpClient, target.siteAbsoluteUrl, {
    step,
    status,
    detail,
    listTitle: target.listTitle || target.listGuid,
    listGuid: target.listGuid,
    itemId,
    code
  });
}

function releaseFill(): void {
  running = false;
}

async function loadItems(
  target: IListFillTarget,
  config: IConfigDraft,
  field: string
): Promise<{ items: Array<{ Id: number; Created?: string; [name: string]: unknown }>; error: string }> {
  const query = recentItemQuery(config.itemQuery.trim() || buildItemQuery(field, config.formulaColumns));
  const guid = normalizeGuid(target.listGuid);
  const web = target.webAbsoluteUrl.replace(/\/+$/, '');
  const response = await target.spHttpClient.get(
    `${web}/_api/web/lists(guid'${guid}')/items?${query}`,
    SPHttpClient.configurations.v1,
    { headers: { Accept: 'application/json;odata=nometadata' } }
  );
  const text = await response.text();
  if (!response.ok) {
    return { items: [], error: `SharePoint returned HTTP ${response.status}. ${text.slice(0, 300)}` };
  }
  const payload = JSON.parse(text) as { value?: Array<{ Id: number; Created?: string; [name: string]: unknown }> };
  return { items: oldestFirst(payload.value || []), error: '' };
}

async function mergeItem(
  spHttpClient: SPHttpClient,
  url: string,
  body: { [name: string]: string | number }
): Promise<{ ok: boolean; detail: string }> {
  const response = await spHttpClient.post(url, SPHttpClient.configurations.v1, {
    headers: jsonHeaders,
    body: JSON.stringify(body)
  });
  const text = await response.text();
  if (!response.ok) {
    return { ok: false, detail: `SharePoint returned HTTP ${response.status}. ${text.slice(0, 300)}` };
  }
  return { ok: true, detail: '' };
}

function itemUrl(target: IListFillTarget, itemId: number): string {
  const guid = normalizeGuid(target.listGuid);
  const web = target.webAbsoluteUrl.replace(/\/+$/, '');
  return `${web}/_api/web/lists(guid'${guid}')/items(${itemId})`;
}

function configUrl(target: IListFillTarget, itemId: number): string {
  const site = target.siteAbsoluteUrl.replace(/\/+$/, '');
  const title = target.configListTitle.replace(/'/g, "''");
  return `${site}/_api/web/lists/getbytitle('${title}')/items(${itemId})`;
}

function createdDate(value: unknown): Date {
  const date = new Date(textOf(value));
  if (Number.isNaN(date.getTime())) {
    return new Date();
  }
  return date;
}

function textOf(value: unknown): string {
  if (value === undefined || (typeof value === 'object' && !value)) {
    return '';
  }
  return String(value);
}

function recentItemQuery(query: string): string {
  const base = query
    .replace(/&?\$orderby=[^&]*/ig, '')
    .replace(/&?\$top=\d+/ig, '')
    .replace(/^&+/, '');
  return `${base}&$orderby=Id desc&$top=50`;
}

function clickRefresh(): boolean {
  const selectors = [
    'button[data-automationid="refreshCommand"]',
    '[data-automationid="refreshCommand"]',
    'button[data-id="Refresh"]',
    'button[name="Refresh"]',
    'button[aria-label="Refresh"]',
    'button[title="Refresh"]',
    '[role="menuitem"][aria-label="Refresh"]',
    '[role="menuitem"][name="Refresh"]'
  ];
  for (let index = 0; index < selectors.length; index += 1) {
    const node = document.querySelector(selectors[index]);
    if (!node) {
      continue;
    }
    const host = node.closest('button, [role="button"], [role="menuitem"]') || node;
    if (host instanceof HTMLElement) {
      host.click();
      return true;
    }
  }
  return false;
}
