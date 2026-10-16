import { CONFIG_FIELD_DEFINITIONS } from './fields';

export const DEFAULT_CONFIG_LIST_TITLE = 'AutoGenFlowConfiguration';

export type ResetPeriod = 'None' | 'Day' | 'Month' | 'Year';

export const RESET_PERIODS: ResetPeriod[] = ['None', 'Day', 'Month', 'Year'];

export interface IConfigDraft {
  id?: number;
  title: string;
  targetListUrl: string;
  targetListGuid: string;
  targetListTitle: string;
  numberColumnInternalName: string;
  formula: string;
  formulaColumns: string;
  counterToken: string;
  counterWidth: number;
  resetPeriod: ResetPeriod;
  isActive: boolean;
  padLength: number;
  currentCount: number;
  lastResetDate: string;
}

export interface ISharePointConfigItem {
  Id: number;
  Title?: string;
  TargetListUrl?: string;
  TargetListGuid?: string;
  TargetListTitle?: string;
  NumberColumnInternalName?: string;
  Formula?: string;
  FormulaColumns?: string;
  CounterToken?: string;
  CounterWidth?: number | string;
  CurrentCount?: number | string;
  ResetPeriod?: string;
  LastResetDate?: string;
  IsActive?: boolean | string | number;
  PadLength?: number | string;
}

export interface IListLocation {
  webAbsoluteUrl: string;
  serverRelativeUrl: string;
}

const GUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const FIELD_PATTERN = /^[A-Za-z_][A-Za-z0-9_]*$/;

const SELECT_FIELDS = [
  'Id',
  'Title',
  ...CONFIG_FIELD_DEFINITIONS.map((field) => field.internalName)
];

export function emptyDraft(): IConfigDraft {
  return {
    title: '',
    targetListUrl: '',
    targetListGuid: '',
    targetListTitle: '',
    numberColumnInternalName: '',
    formula: 'REQ-{yyyy}{MM}-{counter:4}',
    formulaColumns: '',
    counterToken: '{counter:4}',
    counterWidth: 4,
    resetPeriod: 'None',
    isActive: true,
    padLength: 4,
    currentCount: 0,
    lastResetDate: ''
  };
}

export function normalizeGuid(value: string): string {
  return value.trim().replace(/[{}]/g, '').toLowerCase();
}

export function isActiveValue(value: boolean | string | number | undefined): boolean {
  return value === true || value === 1 || value === '1' || value === 'true' || value === 'Yes';
}

export function wholeCount(value: number | string | undefined): number {
  const parsed = Math.floor(Number(value));
  if (!Number.isFinite(parsed) || parsed < 0) {
    return 0;
  }
  return parsed;
}

export function validateDraft(draft: IConfigDraft): string | undefined {
  if (!draft.title.trim()) {
    return 'Enter a title for this rule.';
  }
  if (!isHttpsUrl(draft.targetListUrl)) {
    return 'The list address must be an https URL.';
  }
  if (!GUID_PATTERN.test(normalizeGuid(draft.targetListGuid))) {
    return 'Open AutoGen Flow from the list so the list id is known.';
  }
  if (!FIELD_PATTERN.test(draft.numberColumnInternalName.trim())) {
    return 'Choose a single line of text column.';
  }
  if (!/\{(counter|seq)(?::\d+)?\}/i.test(draft.formula)) {
    return 'Formula must include {counter} or {counter:n}.';
  }
  if (RESET_PERIODS.indexOf(draft.resetPeriod) === -1) {
    return 'Choose a reset period.';
  }
  if (!Number.isInteger(draft.padLength) || draft.padLength < 0 || draft.padLength > 12) {
    return 'Pad length must be a whole number from 0 to 12.';
  }
  if (!Number.isInteger(draft.currentCount) || draft.currentCount < 0) {
    return 'Current count must be a whole number from 0 upward.';
  }
  return undefined;
}

export function mapConfigRow(item: ISharePointConfigItem): IConfigDraft {
  const padLength = Number(item.PadLength);
  const counterWidth = Number(item.CounterWidth);
  return {
    id: item.Id,
    title: item.Title || '',
    targetListUrl: item.TargetListUrl || '',
    targetListGuid: item.TargetListGuid || '',
    targetListTitle: item.TargetListTitle || '',
    numberColumnInternalName: item.NumberColumnInternalName || '',
    formula: item.Formula || '',
    formulaColumns: item.FormulaColumns || '',
    counterToken: item.CounterToken || '{counter}',
    counterWidth: Number.isInteger(counterWidth) && counterWidth >= 0 ? counterWidth : 0,
    resetPeriod: normalizeResetPeriod(item.ResetPeriod),
    isActive: isActiveValue(item.IsActive),
    padLength: Number.isInteger(padLength) && padLength >= 0 ? padLength : 0,
    currentCount: wholeCount(item.CurrentCount),
    lastResetDate: item.LastResetDate || ''
  };
}

export function buildCreateBody(draft: IConfigDraft): Record<string, string | number | boolean> {
  return {
    ...sharedBody(draft),
    CurrentCount: wholeCount(draft.currentCount)
  };
}

export function buildUpdateBody(draft: IConfigDraft, writeCount: boolean): Record<string, string | number | boolean> {
  const body = sharedBody(draft);
  if (writeCount) {
    body.CurrentCount = wholeCount(draft.currentCount);
  }
  return body;
}

export const PERMISSION_EDIT_LIST_ITEMS = 4;

export function hasPermissionFlag(low: string | number | undefined, flag: number): boolean {
  const value = typeof low === 'number' ? low : parseInt(String(low ?? ''), 10);
  if (!Number.isFinite(value)) {
    return false;
  }
  return (value & flag) === flag;
}

export function findConfigForList(rows: IConfigDraft[], listGuid: string): IConfigDraft | undefined {
  const guid = normalizeGuid(listGuid);
  for (let index = 0; index < rows.length; index += 1) {
    if (normalizeGuid(rows[index].targetListGuid) === guid) {
      return rows[index];
    }
  }
  return undefined;
}

export function draftForCurrentList(listTitle: string, listUrl: string, listGuid: string): IConfigDraft {
  return {
    ...emptyDraft(),
    title: listTitle,
    targetListTitle: listTitle,
    targetListUrl: listUrl,
    targetListGuid: normalizeGuid(listGuid)
  };
}

export function buildListAbsoluteUrl(webAbsoluteUrl: string, serverRelativeUrl: string): string {
  const origin = new URL(webAbsoluteUrl).origin;
  const path = serverRelativeUrl.charAt(0) === '/' ? serverRelativeUrl : `/${serverRelativeUrl}`;
  return `${origin}${path}`;
}

export function isFlowConfigList(listTitle: string, serverRelativeUrl: string, configListTitle: string): boolean {
  const title = listTitle.trim().toLowerCase();
  const expected = configListTitle.trim().toLowerCase();
  if (title === expected || title === 'autogenflowconfiguration') {
    return true;
  }
  return /\/lists\/autogenflowconfiguration$/i.test(serverRelativeUrl);
}

export function escapeListTitle(title: string): string {
  return title.replace(/'/g, "''");
}

export function deriveWebFromListUrl(listUrl: string): IListLocation {
  const url = new URL(listUrl.trim());
  if (url.protocol !== 'https:') {
    throw new Error('The list address must use https.');
  }
  let path = decodeURIComponent(url.pathname);
  path = path.replace(/\/Forms\/AllItems\.aspx$/i, '');
  path = path.replace(/\/AllItems\.aspx$/i, '');
  path = path.replace(/\/+$/, '');
  const listsMatch = path.match(/^(.*)\/lists\/[^/]+$/i);
  const webPath = listsMatch ? (listsMatch[1] || '') : path.slice(0, path.lastIndexOf('/'));
  return {
    webAbsoluteUrl: `${url.origin}${webPath || ''}`,
    serverRelativeUrl: path
  };
}

export function buildListItemsUrl(siteAbsoluteUrl: string, listTitle: string): string {
  return `${listApi(siteAbsoluteUrl, listTitle)}/items?$select=${SELECT_FIELDS.join(',')}&$top=200`;
}

export function buildCreateItemUrl(siteAbsoluteUrl: string, listTitle: string): string {
  return `${listApi(siteAbsoluteUrl, listTitle)}/items`;
}

export function buildUpdateItemUrl(siteAbsoluteUrl: string, listTitle: string, itemId: number): string {
  return `${buildCreateItemUrl(siteAbsoluteUrl, listTitle)}(${itemId})`;
}

export function buildListUrl(siteAbsoluteUrl: string, listTitle: string): string {
  return `${listApi(siteAbsoluteUrl, listTitle)}?$select=Id,Title`;
}

export function buildListPermissionsUrl(siteAbsoluteUrl: string, listTitle: string): string {
  return `${listApi(siteAbsoluteUrl, listTitle)}/EffectiveBasePermissions`;
}

export function buildCreateListUrl(siteAbsoluteUrl: string): string {
  return `${trimSlash(siteAbsoluteUrl)}/_api/web/lists`;
}

export function buildCreateListBody(listTitle: string): { BaseTemplate: number; Title: string; Description: string; Hidden: boolean } {
  return {
    BaseTemplate: 100,
    Title: listTitle,
    Description: 'AutoGen rules for Power Automate. The AutoGen Flow command writes one row per list. The flow reads the row and writes the number.',
    Hidden: false
  };
}

export function buildFieldsUrl(siteAbsoluteUrl: string, listTitle: string): string {
  return `${listApi(siteAbsoluteUrl, listTitle)}/fields?$select=InternalName&$top=300`;
}

export function buildCreateFieldUrl(siteAbsoluteUrl: string, listTitle: string): string {
  return `${listApi(siteAbsoluteUrl, listTitle)}/fields/CreateFieldAsXml`;
}

export function isExistingColumnError(body: string): boolean {
  const text = body.toLowerCase();
  return text.indexOf('already exist') >= 0
    || text.indexOf('duplicate field') >= 0
    || text.indexOf('already in use') >= 0;
}

export function sharePointError(status: number, body: string, listTitle: string): string {
  if (status === 404) {
    return `The list "${listTitle}" was not found. Open any page on this site so AutoGen Flow can create it, or save the rule again.`;
  }
  const detail = body.trim().slice(0, 300);
  return `SharePoint returned HTTP ${status}. ${detail}`;
}

function sharedBody(draft: IConfigDraft): Record<string, string | number | boolean> {
  return {
    Title: draft.title.trim(),
    TargetListUrl: draft.targetListUrl.trim(),
    TargetListGuid: normalizeGuid(draft.targetListGuid),
    TargetListTitle: draft.targetListTitle.trim(),
    NumberColumnInternalName: draft.numberColumnInternalName.trim(),
    Formula: draft.formula.trim(),
    FormulaColumns: draft.formulaColumns.trim(),
    CounterToken: draft.counterToken.trim(),
    CounterWidth: draft.counterWidth,
    ResetPeriod: draft.resetPeriod,
    IsActive: draft.isActive,
    PadLength: draft.padLength
  };
}

function listApi(siteAbsoluteUrl: string, listTitle: string): string {
  return `${trimSlash(siteAbsoluteUrl)}/_api/web/lists/getbytitle('${escapeListTitle(listTitle)}')`;
}

function trimSlash(value: string): string {
  return value.replace(/\/+$/, '');
}

function isHttpsUrl(value: string): boolean {
  try {
    return new URL(value.trim()).protocol === 'https:';
  } catch {
    return false;
  }
}

function normalizeResetPeriod(value: string | undefined): ResetPeriod {
  if (value === 'Day' || value === 'Month' || value === 'Year' || value === 'None') {
    return value;
  }
  return 'None';
}
