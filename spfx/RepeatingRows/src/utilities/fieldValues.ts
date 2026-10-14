import { IChildRow, IRepeatingColumn, IRepeatingRowsConfig } from '../models/repeatingRows';
import { applyFormulas } from './formulas';

export function relationshipIdName(config: IRepeatingRowsConfig): string {
  return config.relationshipKind === 'number'
    ? config.relationshipField
    : `${config.relationshipField}Id`;
}

export function childQuery(config: IRepeatingRowsConfig): { select: string[]; expand: string[] } {
  const select = ['Id', 'SortOrder', relationshipIdName(config)];
  const expand: string[] = [];
  config.columns.forEach((column) => {
    if (column.type === 'attachment') {
      return;
    }
    if (column.type === 'lookup' || column.type === 'person') {
      select.push(`${column.internalName}Id`);
      select.push(`${column.internalName}/Title`);
      expand.push(column.internalName);
      return;
    }
    select.push(column.internalName);
  });
  return { select, expand };
}

export function rowPayload(
  row: IChildRow,
  config: IRepeatingRowsConfig
): { [key: string]: string | number | boolean } {
  const payload: { [key: string]: string | number | boolean } = {
    SortOrder: row.sortOrder
  };
  payload[relationshipIdName(config)] = row.parentId;
  config.columns.forEach((column) => {
    writeColumn(payload, column, row.values[column.internalName] || '');
  });
  return payload;
}

export function itemToRow(item: { [key: string]: unknown }, config: IRepeatingRowsConfig): IChildRow {
  const values: { [internalName: string]: string } = {};
  config.columns.forEach((column) => {
    values[column.internalName] = readColumn(item, column);
  });
  const id = numberOrUndefined(item.Id !== undefined ? item.Id : item.ID);
  const sort = numberOrUndefined(item.SortOrder);
  const parent = numberOrUndefined(item[relationshipIdName(config)]);
  const etag = readEtag(item);
  return {
    key: id === undefined ? 'row-unsaved' : `item-${id}`,
    id,
    parentId: parent === undefined ? 0 : parent,
    sortOrder: sort === undefined ? 0 : sort,
    etag,
    values: applyFormulas(values, config.columns)
  };
}

function writeColumn(
  payload: { [key: string]: string | number | boolean },
  column: IRepeatingColumn,
  raw: string
): void {
  if (column.formula || column.type === 'attachment') {
    return;
  }
  if (column.type === 'number' || column.type === 'currency') {
    const numeric = Number(raw);
    if (raw.trim() && Number.isFinite(numeric)) {
      payload[column.internalName] = numeric;
    }
    return;
  }
  if (column.type === 'boolean') {
    payload[column.internalName] = raw === 'true' || raw === '1';
    return;
  }
  if (column.type === 'lookup' || column.type === 'person') {
    const idText = raw.split('|')[0];
    const id = Number(idText);
    if (idText && Number.isFinite(id)) {
      payload[`${column.internalName}Id`] = id;
    }
    return;
  }
  if (column.type === 'date') {
    if (raw) {
      payload[column.internalName] = raw;
    }
    return;
  }
  payload[column.internalName] = raw;
}

function readColumn(item: { [key: string]: unknown }, column: IRepeatingColumn): string {
  if (column.type === 'lookup') {
    const id = item[`${column.internalName}Id`];
    return id === undefined || id === null ? '' : String(id);
  }
  if (column.type === 'person') {
    const id = item[`${column.internalName}Id`];
    const person = item[column.internalName];
    const title = person && typeof person === 'object' ? String((person as { Title?: unknown }).Title || '') : '';
    if (id === undefined || id === null || id === '') {
      return title;
    }
    return title ? `${id}|${title}` : String(id);
  }
  if (column.type === 'boolean') {
    return item[column.internalName] === true ? 'true' : 'false';
  }
  if (column.type === 'date') {
    const raw = item[column.internalName];
    return typeof raw === 'string' ? raw.slice(0, 10) : '';
  }
  if (column.type === 'attachment') {
    return '';
  }
  const raw = item[column.internalName];
  return raw === undefined || raw === null ? '' : String(raw);
}

function readEtag(item: { [key: string]: unknown }): string | undefined {
  const candidates = [item['odata.etag'], item['@odata.etag'], item.etag];
  for (let index = 0; index < candidates.length; index += 1) {
    if (typeof candidates[index] === 'string') {
      return candidates[index] as string;
    }
  }
  return undefined;
}

function numberOrUndefined(value: unknown): number | undefined {
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : undefined;
}
