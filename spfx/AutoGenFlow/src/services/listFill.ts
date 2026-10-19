import { SPHttpClient } from '@microsoft/sp-http';
import { findConfigForList, IConfigDraft, normalizeGuid, wholeCount } from './contract';
import { loadConfigRows } from './configListClient';
import { buildItemQuery } from './itemQuery';
import { planNextCode } from './formula';
import { isBlankNumber, oldestFirst, readColumnValues } from './listFillValues';

export { isBlankNumber, readColumnValues };

export interface IListFillTarget {
  spHttpClient: SPHttpClient;
  siteAbsoluteUrl: string;
  webAbsoluteUrl: string;
  configListTitle: string;
  listGuid: string;
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
  try {
    const rows = await loadConfigRows(target.spHttpClient, target.siteAbsoluteUrl, target.configListTitle);
    const config = findConfigForList(rows, target.listGuid);
    if (!config || !config.isActive || config.id === undefined) {
      return 0;
    }
    const field = config.numberColumnInternalName.trim();
    if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(field)) {
      return 0;
    }
    const items = await loadItems(target, config, field);
    let count = wholeCount(config.currentCount);
    let lastReset = config.lastResetDate;
    let wrote = 0;
    for (let index = 0; index < items.length; index += 1) {
      const item = items[index];
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
      const itemSaved = await mergeItem(target.spHttpClient, itemUrl(target, item.Id), { [field]: planned.code });
      if (!itemSaved) {
        break;
      }
      const countSaved = await mergeItem(
        target.spHttpClient,
        configUrl(target, config.id),
        { CurrentCount: planned.sequence, LastResetDate: planned.lastResetDate }
      );
      if (!countSaved) {
        break;
      }
      count = planned.sequence;
      lastReset = planned.lastResetDate;
      wrote += 1;
    }
    if (wrote > 0) {
      clickRefresh();
    }
    return wrote;
  } catch {
    return 0;
  } finally {
    releaseFill();
  }
}

function releaseFill(): void {
  running = false;
}

async function loadItems(target: IListFillTarget, config: IConfigDraft, field: string): Promise<Array<{ Id: number; Created?: string; [name: string]: unknown }>> {
  const query = recentItemQuery(config.itemQuery.trim() || buildItemQuery(field, config.formulaColumns));
  const guid = normalizeGuid(target.listGuid);
  const web = target.webAbsoluteUrl.replace(/\/+$/, '');
  const response = await target.spHttpClient.get(
    `${web}/_api/web/lists(guid'${guid}')/items?${query}`,
    SPHttpClient.configurations.v1,
    { headers: { Accept: 'application/json;odata=nometadata' } }
  );
  if (!response.ok) {
    return [];
  }
  const payload = await response.json() as { value?: Array<{ Id: number; Created?: string; [name: string]: unknown }> };
  return oldestFirst(payload.value || []);
}

async function mergeItem(spHttpClient: SPHttpClient, url: string, body: { [name: string]: string | number }): Promise<boolean> {
  const response = await spHttpClient.post(url, SPHttpClient.configurations.v1, {
    headers: jsonHeaders,
    body: JSON.stringify(body)
  });
  if (!response.ok) {
    return false;
  }
  await response.text();
  return true;
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

function clickRefresh(): void {
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
      return;
    }
  }
}
