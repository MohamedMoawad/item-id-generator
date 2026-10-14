import { IParentField, IParentValues } from '../models/repeatingRows';

const SKIPPED = ['repeatingrows', 'attachments', 'contenttype', 'taxcatchall', 'linktitle', 'linktitlenomenu'];

export function isParentInput(internalName: string): boolean {
  const name = internalName.toLowerCase();
  if (!name || name.charAt(0) === '_') {
    return false;
  }
  return SKIPPED.indexOf(name) < 0;
}

export function parentFieldType(typeAsString: string): IParentField['type'] | undefined {
  const kind = typeAsString.toLowerCase();
  if (kind === 'text') {
    return 'text';
  }
  if (kind === 'note') {
    return 'note';
  }
  if (kind === 'number') {
    return 'number';
  }
  if (kind === 'currency') {
    return 'currency';
  }
  if (kind === 'choice') {
    return 'choice';
  }
  if (kind === 'datetime') {
    return 'date';
  }
  if (kind === 'boolean') {
    return 'boolean';
  }
  if (kind === 'user') {
    return 'person';
  }
  return undefined;
}

export function parentPayload(fields: IParentField[], values: IParentValues): IParentValues {
  const payload: IParentValues = {};
  fields.forEach((field) => {
    const raw = values[field.internalName];
    const text = raw === undefined || raw === null ? '' : String(raw);
    if (field.type === 'number' || field.type === 'currency') {
      const numeric = Number(text);
      if (text.trim() && Number.isFinite(numeric)) {
        payload[field.internalName] = numeric;
      }
      return;
    }
    if (field.type === 'boolean') {
      payload[field.internalName] = text === 'true' || text === '1';
      return;
    }
    if (field.type === 'person') {
      const id = Number(text.split('|')[0]);
      if (text.indexOf('|') >= 0 && Number.isFinite(id)) {
        payload[`${field.internalName}Id`] = id;
      }
      return;
    }
    if (text) {
      payload[field.internalName] = text;
    }
  });
  return payload;
}

export function validateParent(fields: IParentField[], values: IParentValues): string {
  for (let index = 0; index < fields.length; index += 1) {
    const field = fields[index];
    if (!field.required) {
      continue;
    }
    const raw = values[field.internalName];
    if (raw === undefined || String(raw).trim() === '') {
      return `${field.title} is required.`;
    }
  }
  return '';
}

export function readParentValues(item: { [key: string]: unknown }, fields: IParentField[]): IParentValues {
  const values: IParentValues = {};
  fields.forEach((field) => {
    if (field.type === 'person') {
      const id = item[`${field.internalName}Id`];
      const person = item[field.internalName];
      const title = person && typeof person === 'object' ? String((person as { Title?: unknown }).Title || '') : '';
      values[field.internalName] = id ? `${id}|${title}` : title;
      return;
    }
    if (field.type === 'boolean') {
      values[field.internalName] = item[field.internalName] === true ? 'true' : 'false';
      return;
    }
    if (field.type === 'date') {
      const raw = item[field.internalName];
      values[field.internalName] = typeof raw === 'string' ? raw.slice(0, 10) : '';
      return;
    }
    const raw = item[field.internalName];
    values[field.internalName] = raw === undefined || raw === null ? '' : String(raw);
  });
  return values;
}
