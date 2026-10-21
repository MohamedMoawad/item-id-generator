export function isBlankNumber(value: unknown): boolean {
  if (value === undefined || (typeof value === 'object' && !value)) {
    return true;
  }
  return String(value).trim().length === 0;
}

export function readColumnValues(item: { [name: string]: unknown }, tokens: string[]): { [token: string]: string } {
  const values: { [token: string]: string } = {};
  tokens.forEach((token) => {
    const name = token.split('.')[0];
    const raw = item[name];
    values[token] = token.indexOf('.') >= 0 ? lookupText(raw) : textOf(raw);
  });
  return values;
}

export function oldestFirst<T>(items: T[]): T[] {
  return items.slice().reverse();
}

function lookupText(value: unknown): string {
  if (Array.isArray(value)) {
    return value.map((item) => lookupText(item)).filter((text) => text.length > 0).join(', ');
  }
  if (value && typeof value === 'object') {
    const record = value as { [key: string]: unknown };
    if (Array.isArray(record.results)) {
      return lookupText(record.results);
    }
    const picked = firstText(record, ['Title', 'LookupValue', 'EMail', 'Email', 'Value']);
    if (picked !== undefined) {
      return textOf(picked);
    }
  }
  return textOf(value);
}

function firstText(record: { [key: string]: unknown }, names: string[]): unknown {
  for (let index = 0; index < names.length; index += 1) {
    const picked = record[names[index]];
    if (picked !== undefined && picked !== null && typeof picked !== 'object') {
      return picked;
    }
  }
  return undefined;
}

function textOf(value: unknown): string {
  if (value === undefined || (typeof value === 'object' && !value)) {
    return '';
  }
  return String(value);
}
