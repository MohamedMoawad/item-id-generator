import { CONFIG_FIELD_DEFINITIONS } from './requestNumberFields';

export const DEFAULT_CONFIG_LIST_TITLE = 'AutoGenFeatureConfiguration';

export type ResetPeriod = 'None' | 'Day' | 'Month' | 'Year';

export const RESET_PERIODS: ResetPeriod[] = ['None', 'Day', 'Month', 'Year'];

export interface IConfigDraft {
  id?: number;
  etag?: string;
  title: string;
  targetListUrl: string;
  targetListGuid: string;
  targetListTitle?: string;
  loadedTargetListGuid?: string;
  numberColumnInternalName: string;
  formula: string;
  resetPeriod: ResetPeriod;
  isActive: boolean;
  padLength: number;
  currentCount?: number;
  lastResetDate?: string;
  webhookSubscriptionId?: string;
  modified?: string;
  modifiedBy?: string;
}

export interface ISharePointConfigItem {
  Id: number;
  Title?: string;
  TargetListUrl?: string;
  TargetListGuid?: string;
  NumberColumnInternalName?: string;
  Formula?: string;
  CurrentCount?: number;
  ResetPeriod?: string;
  LastResetDate?: string;
  IsActive?: boolean | string | number;
  PadLength?: number | string;
  WebhookSubscriptionId?: string;
  Modified?: string;
  Editor?: { Title?: string };
  'odata.etag'?: string;
  '@odata.etag'?: string;
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
    numberColumnInternalName: '',
    formula: 'REQ-{yyyy}{MM}-{counter}',
    resetPeriod: 'Month',
    isActive: true,
    padLength: 4
  };
}

export function normalizeGuid(value: string): string {
  return value.trim().replace(/[{}]/g, '').toLowerCase();
}

export function isActiveValue(value: boolean | string | number | undefined): boolean {
  return value === true || value === 1 || value === '1' || value === 'true' || value === 'Yes';
}

export function isPlaceholderSetting(value: string): boolean {
  const text = value.trim();
  return text.length === 0 || text.indexOf('<') >= 0;
}

export function validateDraft(draft: IConfigDraft): string | undefined {
  if (!draft.title.trim()) {
    return 'Enter a title for this rule.';
  }
  if (!isHttpsUrl(draft.targetListUrl)) {
    return 'Target list URL must be an https URL.';
  }
  if (!GUID_PATTERN.test(normalizeGuid(draft.targetListGuid))) {
    return 'Paste the list URL and choose Resolve list before saving.';
  }
  if (!draft.numberColumnInternalName.trim()) {
    return 'Choose a single line of text column.';
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
  return undefined;
}

export function mapConfigRow(item: ISharePointConfigItem): IConfigDraft {
  const guid = item.TargetListGuid || '';
  const padLength = Number(item.PadLength);
  return {
    id: item.Id,
    etag: item['odata.etag'] || item['@odata.etag'] || '*',
    title: item.Title || '',
    targetListUrl: item.TargetListUrl || '',
    targetListGuid: guid,
    loadedTargetListGuid: guid,
    numberColumnInternalName: item.NumberColumnInternalName || '',
    formula: item.Formula || '',
    resetPeriod: normalizeResetPeriod(item.ResetPeriod),
    isActive: isActiveValue(item.IsActive),
    padLength: Number.isInteger(padLength) && padLength >= 0 ? padLength : 0,
    currentCount: Number(item.CurrentCount) || 0,
    lastResetDate: item.LastResetDate || '',
    webhookSubscriptionId: item.WebhookSubscriptionId || '',
    modified: item.Modified || '',
    modifiedBy: item.Editor && item.Editor.Title ? item.Editor.Title : ''
  };
}

export function buildCreateBody(draft: IConfigDraft): Record<string, string | number | boolean> {
  return {
    Title: draft.title.trim(),
    TargetListUrl: draft.targetListUrl.trim(),
    TargetListGuid: normalizeGuid(draft.targetListGuid),
    NumberColumnInternalName: draft.numberColumnInternalName.trim(),
    Formula: draft.formula.trim(),
    ResetPeriod: draft.resetPeriod,
    IsActive: draft.isActive,
    PadLength: draft.padLength,
    CurrentCount: 0
  };
}

export function buildUpdateBody(draft: IConfigDraft): Record<string, string | number | boolean> {
  const body: Record<string, string | number | boolean> = {
    Title: draft.title.trim(),
    TargetListUrl: draft.targetListUrl.trim(),
    TargetListGuid: normalizeGuid(draft.targetListGuid),
    NumberColumnInternalName: draft.numberColumnInternalName.trim(),
    Formula: draft.formula.trim(),
    ResetPeriod: draft.resetPeriod,
    IsActive: draft.isActive,
    PadLength: draft.padLength
  };
  const previous = draft.loadedTargetListGuid ? normalizeGuid(draft.loadedTargetListGuid) : '';
  if (previous && previous !== normalizeGuid(draft.targetListGuid)) {
    body.WebhookSubscriptionId = '';
  }
  return body;
}

export const PERMISSION_EDIT_LIST_ITEMS = 4;

export const REGISTER_WEBHOOK_PROPERTY = 'RequestNumberRegisterWebhookUrl';

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
    targetListUrl: listUrl,
    targetListGuid: normalizeGuid(listGuid)
  };
}

export function buildListAbsoluteUrl(webAbsoluteUrl: string, serverRelativeUrl: string): string {
  const origin = new URL(webAbsoluteUrl).origin;
  const path = serverRelativeUrl.charAt(0) === '/' ? serverRelativeUrl : `/${serverRelativeUrl}`;
  return `${origin}${path}`;
}

export function isConfigList(listTitle: string, serverRelativeUrl: string, configListTitle: string): boolean {
  const title = listTitle.trim().toLowerCase();
  if (title === configListTitle.trim().toLowerCase() || title === 'autogenfeatureconfiguration' || title === 'requestnumberconfig') {
    return true;
  }
  return /\/lists\/(autogenfeatureconfiguration|requestnumberconfig)$/i.test(serverRelativeUrl);
}

export function buildRegisterWebhookBody(configSiteUrl: string, configItemId: number): {
  configSiteUrl: string;
  configItemId: number;
} {
  return {
    configSiteUrl: trimSlash(configSiteUrl),
    configItemId
  };
}

export function escapeListTitle(title: string): string {
  return title.replace(/'/g, "''");
}

export function deriveWebFromListUrl(listUrl: string): IListLocation {
  const url = new URL(listUrl.trim());
  if (url.protocol !== 'https:') {
    throw new Error('Target list URL must use https.');
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
  return `${listApi(siteAbsoluteUrl, listTitle)}/items?$select=${SELECT_FIELDS.join(',')},Modified,Editor/Title&$expand=Editor&$top=200`;
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

export function buildAllPropertiesUrl(siteAbsoluteUrl: string): string {
  return `${trimSlash(siteAbsoluteUrl)}/_api/web/AllProperties`;
}

export function buildStorageEntityUrl(siteAbsoluteUrl: string): string {
  return `${trimSlash(siteAbsoluteUrl)}/_api/web/GetStorageEntity('${REGISTER_WEBHOOK_PROPERTY}')`;
}

export function buildTenantSettingsUrl(siteAbsoluteUrl: string): string {
  return `${trimSlash(siteAbsoluteUrl)}/_api/SP_TenantSettings_Current`;
}

export function readStorageEntityValue(payload: unknown): string {
  const record = asRecord(payload);
  if (record['odata.null'] === true) {
    return '';
  }
  const direct = textValue(record.Value);
  if (direct) {
    return direct;
  }
  const nested = asRecord(record.d);
  const method = asRecord(nested.GetStorageEntity);
  return textValue(method.Value) || textValue(nested.Value);
}

export function readCorporateCatalogUrl(payload: unknown): string {
  const record = asRecord(payload);
  const nested = asRecord(record.d);
  const url = textValue(record.CorporateCatalogUrl) || textValue(nested.CorporateCatalogUrl);
  if (!url || url.indexOf('https://') !== 0) {
    return '';
  }
  return trimSlash(url);
}

export function buildSetStorageEntity(catalogSiteUrl: string, value: string): { url: string; headers: { [key: string]: string }; body: string } {
  return {
    url: `${trimSlash(catalogSiteUrl)}/_api/web/SetStorageEntity`,
    headers: {
      Accept: 'application/json;odata=verbose',
      'Content-Type': 'application/json;odata=verbose',
      'OData-Version': '3.0'
    },
    body: JSON.stringify({
      entity: {
        __metadata: { type: 'SP.AppKeyValue' },
        Key: REGISTER_WEBHOOK_PROPERTY,
        Value: value,
        Description: 'Organization RegisterWebhook URL for request numbers',
        Comment: 'Used by every site collection where the app is added'
      }
    })
  };
}

function asRecord(value: unknown): { [key: string]: unknown } {
  return value && typeof value === 'object' ? value as { [key: string]: unknown } : {};
}

function textValue(value: unknown): string {
  return typeof value === 'string' ? value.trim() : '';
}

// SP.PropertyValues accepts MERGE only under OData 3. SPHttpClient v1 defaults to
// OData 4, which turns MERGE into PATCH and SharePoint returns HTTP 400.
export function buildPropertyBagMerge(propertyName: string, value: string): { headers: { [key: string]: string }; body: string } {
  const payload: { [key: string]: string | { type: string } } = {
    __metadata: { type: 'SP.PropertyValues' }
  };
  payload[propertyName] = value;
  return {
    headers: {
      Accept: 'application/json;odata=verbose',
      'Content-Type': 'application/json;odata=verbose',
      'OData-Version': '3.0',
      'IF-MATCH': '*',
      'X-HTTP-Method': 'MERGE'
    },
    body: JSON.stringify(payload)
  };
}

export function buildCreateListUrl(siteAbsoluteUrl: string): string {
  return `${trimSlash(siteAbsoluteUrl)}/_api/web/lists`;
}

export function buildCreateListBody(listTitle: string): { BaseTemplate: number; Title: string; Description: string; Hidden: boolean } {
  return {
    BaseTemplate: 100,
    Title: listTitle,
    Description: 'Per-list Autogen settings. Open Autogen Feature on a list to view or edit them.',
    Hidden: false
  };
}

export function buildFieldsUrl(siteAbsoluteUrl: string, listTitle: string): string {
  return `${listApi(siteAbsoluteUrl, listTitle)}/fields?$select=InternalName&$top=200`;
}

export function buildCreateFieldUrl(siteAbsoluteUrl: string, listTitle: string): string {
  return `${listApi(siteAbsoluteUrl, listTitle)}/fields/CreateFieldAsXml`;
}

export function buildGetListUrl(listUrl: string): string {
  const location = deriveWebFromListUrl(listUrl);
  const quoted = `'${location.serverRelativeUrl.replace(/'/g, "''")}'`;
  return `${location.webAbsoluteUrl}/_api/web/GetList(@listUrl)?@listUrl=${encodeURIComponent(quoted)}&$select=Id,Title`;
}

export function sharePointError(status: number, body: string, listTitle: string): string {
  if (status === 404) {
    return `The list "${listTitle}" was not found on this site collection. Choose Ensure ${listTitle} to create it.`;
  }
  const detail = body.trim().slice(0, 300);
  return `SharePoint returned HTTP ${status}. ${detail}`;
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
