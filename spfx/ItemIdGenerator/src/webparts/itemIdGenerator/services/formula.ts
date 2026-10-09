import { ResetPeriod } from './configContract';

export function periodKeyFor(resetPeriod: string, date: Date): string {
  const period = normalizePeriod(resetPeriod);
  const utc = new Date(date.getTime());
  if (period === 'day') {
    return formatUtc(utc, 'yyyy-MM-dd');
  }
  if (period === 'month') {
    return formatUtc(utc, 'yyyy-MM');
  }
  if (period === 'year') {
    return formatUtc(utc, 'yyyy');
  }
  return '';
}

export function periodRolledOver(resetPeriod: string, lastResetDate: string | undefined, now: Date): boolean {
  const period = normalizePeriod(resetPeriod);
  if (period === 'none' || !lastResetDate || !lastResetDate.trim()) {
    return false;
  }
  const last = new Date(lastResetDate);
  if (Number.isNaN(last.getTime())) {
    return false;
  }
  return periodKeyFor(period, last) !== periodKeyFor(period, now);
}

export function applyFormula(
  formula: string,
  sequence: number,
  date: Date,
  padLength: number,
  columns?: { [token: string]: string }
): string {
  if (!formula || !hasCounterToken(formula)) {
    throw new Error('Formula must include {counter} or {counter:n} so each item gets a distinct code.');
  }
  if (padLength < 0 || padLength > 12) {
    throw new Error('PadLength must be from 0 to 12.');
  }
  const utc = new Date(date.getTime());
  const values: { [token: string]: string } = {
    yyyy: formatUtc(utc, 'yyyy'),
    yy: formatUtc(utc, 'yy'),
    MM: formatUtc(utc, 'MM'),
    dd: formatUtc(utc, 'dd'),
    HH: formatUtc(utc, 'HH'),
    mm: formatUtc(utc, 'mm')
  };
  return formula.replace(/\{([^{}]+)\}/g, (match: string, raw: string) => {
    const token = raw.trim();
    const counter = /^(counter|seq)(?::(\d+))?$/i.exec(token);
    if (counter) {
      if (!counter[2]) {
        return pad(sequence, padLength);
      }
      const width = Number(counter[2]);
      if (!Number.isInteger(width) || width < 1 || width > 12) {
        throw new Error('Counter width must be from 1 to 12.');
      }
      return pad(sequence, width);
    }
    if (Object.prototype.hasOwnProperty.call(values, token)) {
      return values[token];
    }
    if (!columns) {
      return match;
    }
    return columnText(columns, token);
  });
}

export function planNextCode(input: {
  formula: string;
  resetPeriod: ResetPeriod | string;
  lastResetDate?: string;
  currentCount: number;
  padLength: number;
  now: Date;
  columns?: { [token: string]: string };
}): { code: string; sequence: number; lastResetDate: string } {
  const rolled = periodRolledOver(input.resetPeriod, input.lastResetDate, input.now);
  const sequence = (rolled ? 0 : input.currentCount) + 1;
  return {
    code: applyFormula(input.formula, sequence, input.now, input.padLength, input.columns),
    sequence,
    lastResetDate: input.now.toISOString()
  };
}

export interface IFormulaColumn {
  internalName: string;
  title: string;
  type: string;
  token: string;
}

export function hasCounterToken(formula: string): boolean {
  return /\{(counter|seq)(?::\d+)?\}/i.test(formula);
}

export function isLookupLike(type: string): boolean {
  const text = (type || '').toLowerCase();
  return text.indexOf('lookup') >= 0 || text === 'user' || text === 'usermulti';
}

export function columnToken(internalName: string, type: string): string {
  return isLookupLike(type) ? `{${internalName}.title}` : `{${internalName}}`;
}

export function columnNamesInFormula(formula: string): string[] {
  const names: string[] = [];
  const pattern = /\{([^{}]+)\}/g;
  let match = pattern.exec(formula);
  while (match) {
    const token = match[1].trim();
    match = pattern.exec(formula);
    if (/^(counter|seq)(?::\d+)?$/i.test(token) || /^(yyyy|yy|MM|dd|HH|mm)$/.test(token)) {
      continue;
    }
    const name = token.split('.')[0];
    if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(name) && names.indexOf(name) === -1) {
      names.push(name);
    }
  }
  return names;
}

const SKIP_FIELD_TYPES: { [type: string]: boolean } = {
  Computed: true,
  Invalid: true,
  Attachments: true,
  WorkflowStatus: true,
  Geolocation: true,
  Thumbnail: true
};

const SKIP_FIELD_NAMES: { [name: string]: boolean } = {
  Attachments: true,
  Edit: true,
  LinkTitle: true,
  LinkTitleNoMenu: true,
  DocIcon: true,
  AppAuthor: true,
  AppEditor: true,
  FolderChildCount: true,
  ItemChildCount: true,
  ContentType: true,
  ComplianceAssetId: true,
  _UIVersionString: true,
  FileRef: true,
  FileDirRef: true,
  FSObjType: true,
  UniqueId: true,
  GUID: true,
  MetaInfo: true,
  PermMask: true
};

export function formulaColumns(fields: Array<{ InternalName?: string; Title?: string; TypeAsString?: string; Hidden?: boolean | string }>): IFormulaColumn[] {
  return fields
    .filter((field) => {
      const name = field.InternalName || '';
      return !!name &&
        field.Hidden !== true &&
        field.Hidden !== 'True' &&
        !SKIP_FIELD_TYPES[field.TypeAsString || ''] &&
        !SKIP_FIELD_NAMES[name] &&
        /^[A-Za-z_][A-Za-z0-9_]*$/.test(name);
    })
    .map((field) => ({
      internalName: field.InternalName || '',
      title: field.Title || field.InternalName || '',
      type: field.TypeAsString || 'Text',
      token: columnToken(field.InternalName || '', field.TypeAsString || 'Text')
    }))
    .sort((left, right) => left.title.localeCompare(right.title));
}

export function itemQueryParts(
  numberField: string,
  names: string[],
  columns: IFormulaColumn[]
): { select: string[]; expand: string[] } {
  const select = ['Id', numberField];
  const expand: string[] = [];
  names.forEach((name) => {
    if (name === 'Id' || name.toLowerCase() === numberField.toLowerCase()) {
      return;
    }
    const column = columns.filter((item) => item.internalName === name)[0];
    if (column && isLookupLike(column.type)) {
      select.push(`${name}/Title`, `${name}/Id`);
      expand.push(name);
      return;
    }
    select.push(name);
  });
  return { select: unique(select), expand: unique(expand) };
}

export function columnMapForItem(
  item: { [key: string]: unknown },
  names: string[],
  columns: IFormulaColumn[]
): { [token: string]: string } {
  const map: { [token: string]: string } = {};
  names.forEach((name) => {
    const column = columns.filter((itemColumn) => itemColumn.internalName === name)[0];
    const value = item[name];
    const lookup = !column || isLookupLike(column.type) || (value !== null && typeof value === 'object');
    const title = textFromFieldValue(value, 'title');
    const plain = textFromFieldValue(value);
    map[name] = lookup ? (title || plain) : plain;
    map[`${name}.title`] = title || plain;
    map[`${name}.Title`] = map[`${name}.title`];
    map[`${name}.id`] = textFromFieldValue(value, 'id');
    map[`${name}.Id`] = map[`${name}.id`];
  });
  return map;
}

export function textFromFieldValue(value: unknown, property?: string): string {
  if (value === null || value === undefined) {
    return '';
  }
  if (Array.isArray(value)) {
    return value
      .map((item) => textFromFieldValue(item, property))
      .filter((text) => text.length > 0)
      .join(', ');
  }
  if (typeof value === 'object') {
    const record = value as { [key: string]: unknown };
    if (Array.isArray(record.results)) {
      return textFromFieldValue(record.results, property);
    }
    if (property) {
      const picked = readProperty(record, property);
      if (picked !== undefined && picked !== null && typeof picked !== 'object') {
        return formatScalar(picked);
      }
      if (picked !== undefined && picked !== null) {
        return textFromFieldValue(picked);
      }
      return '';
    }
    const title = readProperty(record, 'Title') ?? readProperty(record, 'LookupValue') ?? readProperty(record, 'Label') ?? readProperty(record, 'EMail') ?? readProperty(record, 'Email');
    if (title !== undefined && title !== null && typeof title !== 'object') {
      return formatScalar(title);
    }
    const description = readProperty(record, 'Description');
    const url = readProperty(record, 'Url');
    if (typeof description === 'string' && description.trim()) {
      return description.trim();
    }
    if (typeof url === 'string') {
      return url.trim();
    }
    return '';
  }
  return formatScalar(value);
}

export function isBlankNumber(value: unknown): boolean {
  if (value === null || value === undefined) {
    return true;
  }
  return String(value).trim().length === 0;
}

function columnText(columns: { [token: string]: string }, token: string): string {
  if (Object.prototype.hasOwnProperty.call(columns, token)) {
    return columns[token];
  }
  const wanted = token.toLowerCase();
  const names = Object.keys(columns);
  for (let index = 0; index < names.length; index += 1) {
    if (names[index].toLowerCase() === wanted) {
      return columns[names[index]];
    }
  }
  return '';
}

function readProperty(record: { [key: string]: unknown }, property: string): unknown {
  if (Object.prototype.hasOwnProperty.call(record, property)) {
    return record[property];
  }
  const wanted = property.toLowerCase();
  const names = Object.keys(record);
  for (let index = 0; index < names.length; index += 1) {
    if (names[index].toLowerCase() === wanted) {
      return record[names[index]];
    }
  }
  return undefined;
}

function formatScalar(value: unknown): string {
  if (typeof value === 'boolean') {
    return value ? 'Yes' : 'No';
  }
  if (typeof value === 'number') {
    return String(value);
  }
  const text = String(value).trim();
  const date = /^(\d{4}-\d{2}-\d{2})T/.exec(text);
  return date ? date[1] : text;
}

function unique(values: string[]): string[] {
  const seen: { [value: string]: boolean } = {};
  return values.filter((value) => {
    if (seen[value]) {
      return false;
    }
    seen[value] = true;
    return true;
  });
}

function normalizePeriod(value: string): 'none' | 'day' | 'month' | 'year' {
  const text = (value || '').trim().toLowerCase();
  if (text === 'day' || text === 'month' || text === 'year') {
    return text;
  }
  return 'none';
}

function pad(sequence: number, width: number): string {
  let text = String(sequence);
  while (text.length < width) {
    text = `0${text}`;
  }
  return text;
}

function formatUtc(date: Date, pattern: string): string {
  const year = date.getUTCFullYear();
  const month = date.getUTCMonth() + 1;
  const day = date.getUTCDate();
  const hour = date.getUTCHours();
  const minute = date.getUTCMinutes();
  return pattern
    .replace('yyyy', String(year))
    .replace('yy', String(year).slice(-2))
    .replace('MM', two(month))
    .replace('dd', two(day))
    .replace('HH', two(hour))
    .replace('mm', two(minute));
}

function two(value: number): string {
  return value < 10 ? `0${value}` : String(value);
}
