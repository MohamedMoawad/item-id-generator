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
    if (token.indexOf('.') >= 0 && raw && typeof raw === 'object') {
      const record = raw as { Title?: unknown; Value?: unknown };
      values[token] = textOf(record.Title !== undefined ? record.Title : record.Value);
      return;
    }
    values[token] = textOf(raw);
  });
  return values;
}

export function oldestFirst<T>(items: T[]): T[] {
  return items.slice().reverse();
}

function textOf(value: unknown): string {
  if (value === undefined || (typeof value === 'object' && !value)) {
    return '';
  }
  return String(value);
}
