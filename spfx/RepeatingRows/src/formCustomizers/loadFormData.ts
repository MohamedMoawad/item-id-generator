import { IChildRow, IParentField, IParentValues, IRepeatingRowsConfig } from '../models/repeatingRows';
import { HttpChildRepository } from '../repositories/httpChildRepository';
import { HttpParentRepository } from '../repositories/httpParentRepository';
import { ISharePointGateway, assertOk, readRows } from '../services/httpTypes';
import { loadLookupOptions } from '../services/lookupLoader';
import { LookupCache } from '../utilities/lookupCache';
import { isParentInput, parentFieldType, readParentValues } from '../utilities/parentValues';

export interface ILoadedForm {
  fields: IParentField[];
  values: IParentValues;
  etag?: string;
  rows: IChildRow[];
  hasMore: boolean;
  lookups: { [internalName: string]: ILookupOptionMap };
}

type ILookupOptionMap = { id: string; text: string; group?: string }[];

const cache = new LookupCache(5 * 60 * 1000);

export async function loadFormData(
  gateway: ISharePointGateway,
  listId: string,
  itemId: number | undefined,
  config: IRepeatingRowsConfig
): Promise<ILoadedForm> {
  const fields = await loadParentFields(gateway, listId);
  const lookups = await loadLookupOptions(gateway, config, cache, Date.now());
  if (!itemId) {
    return { fields, values: {}, rows: [], hasMore: false, lookups };
  }
  const select: string[] = [];
  const expand: string[] = [];
  fields.forEach((field) => {
    if (field.type === 'person') {
      select.push(`${field.internalName}Id`, `${field.internalName}/Title`);
      expand.push(field.internalName);
      return;
    }
    select.push(field.internalName);
  });
  const parent = new HttpParentRepository(gateway, listId, select, expand);
  const record = await parent.load(itemId);
  const children = new HttpChildRepository(gateway, config);
  const page = await children.loadPage(itemId, 0);
  return {
    fields,
    values: readParentValues(record.item, fields),
    etag: record.etag,
    rows: page.rows,
    hasMore: page.hasMore,
    lookups
  };
}

export async function loadParentFields(gateway: ISharePointGateway, listId: string): Promise<IParentField[]> {
  const result = await gateway.get(
    `${gateway.webUrl}/_api/web/lists(guid'${listId}')/fields?$select=InternalName,Title,TypeAsString,Required,Choices&$filter=Hidden eq false and ReadOnlyField eq false`
  );
  assertOk(result);
  const rows = readRows<{
    InternalName?: string;
    Title?: string;
    TypeAsString?: string;
    Required?: boolean;
    Choices?: unknown;
  }>(result.text);
  const fields: IParentField[] = [];
  rows.forEach((row) => {
    const internalName = row.InternalName || '';
    const type = parentFieldType(row.TypeAsString || '');
    if (!type || !isParentInput(internalName)) {
      return;
    }
    const field: IParentField = {
      internalName,
      title: row.Title || internalName,
      type,
      required: row.Required === true
    };
    const choices = readChoices(row.Choices);
    if (choices.length > 0) {
      field.choices = choices;
    }
    fields.push(field);
  });
  return fields;
}

function readChoices(value: unknown): string[] {
  if (Array.isArray(value)) {
    return value.map((choice) => String(choice));
  }
  if (value && typeof value === 'object') {
    const results = (value as { results?: unknown[] }).results;
    if (Array.isArray(results)) {
      return results.map((choice) => String(choice));
    }
  }
  return [];
}
