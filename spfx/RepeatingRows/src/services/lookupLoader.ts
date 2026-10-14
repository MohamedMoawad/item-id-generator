import { ILookupOption, IRepeatingColumn, IRepeatingRowsConfig } from '../models/repeatingRows';
import { LookupCache } from '../utilities/lookupCache';
import { assertOk, ISharePointGateway, odataTitle, readRows } from './httpTypes';

export async function loadLookupOptions(
  gateway: ISharePointGateway,
  config: IRepeatingRowsConfig,
  cache: LookupCache,
  now: number
): Promise<{ [internalName: string]: ILookupOption[] }> {
  const options: { [internalName: string]: ILookupOption[] } = {};
  const columns = config.columns.filter((column) => column.type === 'lookup' && column.lookupListTitle);
  for (let index = 0; index < columns.length; index += 1) {
    options[columns[index].internalName] = await loadColumn(gateway, columns[index], cache, now);
  }
  return options;
}

async function loadColumn(
  gateway: ISharePointGateway,
  column: IRepeatingColumn,
  cache: LookupCache,
  now: number
): Promise<ILookupOption[]> {
  const textField = column.lookupTextField || 'Title';
  const groupField = column.lookupGroupField || '';
  const key = `${column.lookupListTitle}|${textField}|${groupField}`;
  const cached = cache.get(key, now);
  if (cached) {
    return cached;
  }
  const select = groupField ? `Id,${textField},${groupField}` : `Id,${textField}`;
  const url = `${gateway.webUrl}/_api/web/lists/getbytitle('${odataTitle(column.lookupListTitle || '')}')/items?$select=${select}&$top=500`;
  const result = await gateway.get(url);
  assertOk(result);
  const options = readRows<{ [key: string]: unknown }>(result.text).map((item) => ({
    id: String(item.Id || ''),
    text: String(item[textField] || ''),
    group: groupField ? String(item[groupField] || '') : undefined
  }));
  cache.set(key, options, now);
  return options;
}
