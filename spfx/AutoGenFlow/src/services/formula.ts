import { ResetPeriod } from './contract';

export interface IPreparedRule {
  formula: string;
  counterToken: string;
  counterWidth: number;
  formulaColumns: string;
}

export interface IFormulaColumn {
  internalName: string;
  title: string;
  type: string;
  token: string;
}

const DATE_TOKENS: { [token: string]: boolean } = {
  yyyy: true,
  yy: true,
  MM: true,
  dd: true,
  HH: true,
  mm: true
};

export function periodKeyFor(resetPeriod: string, date: Date): string {
  const period = normalizePeriod(resetPeriod);
  if (period === 'day') {
    return formatUtc(date, 'yyyy-MM-dd');
  }
  if (period === 'month') {
    return formatUtc(date, 'yyyy-MM');
  }
  if (period === 'year') {
    return formatUtc(date, 'yyyy');
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

export function normalizeFormula(formula: string): string {
  let text = formula.replace(/[-_/.\s]*\{id\}/gi, '');
  for (let pass = 0; pass < 4; pass += 1) {
    const next = text.replace(/(\{(?:counter|seq)(?::\d+)?\})\s*(\{(?:counter|seq)(?::\d+)?\})/gi, '$1');
    if (next === text) {
      break;
    }
    text = next;
  }
  return text.trim();
}

export function prepareRule(formula: string, padLength: number): IPreparedRule {
  const normalized = normalizeFormula(formula);
  return {
    formula: normalized,
    counterToken: firstCounterToken(normalized),
    counterWidth: counterWidth(normalized, padLength),
    formulaColumns: columnTokensInFormula(normalized).join(',')
  };
}

export function firstCounterToken(formula: string): string {
  const match = /\{((?:counter|seq)(?::\d+)?)\}/i.exec(formula);
  if (!match) {
    return '{counter}';
  }
  return `{${match[1]}}`;
}

export function counterWidth(formula: string, padLength: number): number {
  const match = /\{(?:counter|seq)(?::(\d+))?\}/i.exec(formula);
  if (match && match[1]) {
    const width = Number(match[1]);
    if (Number.isInteger(width) && width >= 1 && width <= 12) {
      return width;
    }
  }
  return clampPad(padLength);
}

export function columnTokensInFormula(formula: string): string[] {
  const tokens: string[] = [];
  const pattern = /\{([^{}]+)\}/g;
  let match = pattern.exec(formula);
  while (match) {
    const token = match[1].trim();
    match = pattern.exec(formula);
    if (isReservedToken(token)) {
      continue;
    }
    const name = token.split('.')[0];
    if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(name) && tokens.indexOf(token) === -1) {
      tokens.push(token);
    }
  }
  return tokens;
}

export function applyFormula(
  formula: string,
  sequence: number,
  date: Date,
  padLength: number,
  columns?: { [token: string]: string }
): string {
  const source = normalizeFormula(formula);
  if (!source || !hasCounterToken(source)) {
    throw new Error('Formula must include {counter} or {counter:n} so each item gets a distinct code.');
  }
  const width = clampPad(padLength);
  const values: { [token: string]: string } = {
    yyyy: formatUtc(date, 'yyyy'),
    yy: formatUtc(date, 'yy'),
    MM: formatUtc(date, 'MM'),
    dd: formatUtc(date, 'dd'),
    HH: formatUtc(date, 'HH'),
    mm: formatUtc(date, 'mm')
  };
  return source.replace(/\{([^{}]+)\}/g, (match: string, raw: string) => {
    const token = raw.trim();
    const counter = /^(counter|seq)(?::(\d+))?$/i.exec(token);
    if (counter) {
      if (!counter[2]) {
        return padSequence(sequence, width);
      }
      const tokenWidth = Number(counter[2]);
      if (!Number.isInteger(tokenWidth) || tokenWidth < 1 || tokenWidth > 12) {
        throw new Error('Counter width must be from 1 to 12.');
      }
      return padSequence(sequence, tokenWidth);
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
  currentCount: number | string;
  padLength: number;
  now: Date;
  columns?: { [token: string]: string };
}): { code: string; sequence: number; lastResetDate: string } {
  const rolled = periodRolledOver(input.resetPeriod, input.lastResetDate, input.now);
  const sequence = (rolled ? 0 : wholeNumber(input.currentCount)) + 1;
  return {
    code: applyFormula(input.formula, sequence, input.now, input.padLength, input.columns),
    sequence,
    lastResetDate: input.now.toISOString()
  };
}

export function padSequence(sequence: number, width: number): string {
  return padForFlow(sequence, width);
}

export function padForFlow(sequence: number, width: number): string {
  const text = String(wholeNumber(sequence));
  const digits = clampPad(width);
  if (digits === 0 || text.length > digits) {
    return text;
  }
  const padded = `000000000000${text}`;
  const start = 12 - (digits - text.length);
  return padded.substring(start, start + digits);
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

export function textColumns(columns: IFormulaColumn[]): IFormulaColumn[] {
  return columns.filter((column) => column.type === 'Text');
}

function isReservedToken(token: string): boolean {
  return /^(counter|seq)(?::\d+)?$/i.test(token) || DATE_TOKENS[token] === true;
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

function normalizePeriod(value: string): 'none' | 'day' | 'month' | 'year' {
  const text = (value || '').trim().toLowerCase();
  if (text === 'day' || text === 'month' || text === 'year') {
    return text;
  }
  return 'none';
}

function wholeNumber(value: number | string | undefined): number {
  const parsed = Math.floor(Number(value));
  if (!Number.isFinite(parsed) || parsed < 0) {
    return 0;
  }
  return parsed;
}

function clampPad(padLength: number): number {
  const width = Math.floor(Number(padLength));
  if (!Number.isInteger(width) || width < 0) {
    return 0;
  }
  if (width > 12) {
    return 12;
  }
  return width;
}

function formatUtc(date: Date, pattern: string): string {
  const year = date.getUTCFullYear();
  return pattern
    .replace('yyyy', String(year))
    .replace('yy', String(year).slice(-2))
    .replace('MM', two(date.getUTCMonth() + 1))
    .replace('dd', two(date.getUTCDate()))
    .replace('HH', two(date.getUTCHours()))
    .replace('mm', two(date.getUTCMinutes()));
}

function two(value: number): string {
  return value < 10 ? `0${value}` : String(value);
}
