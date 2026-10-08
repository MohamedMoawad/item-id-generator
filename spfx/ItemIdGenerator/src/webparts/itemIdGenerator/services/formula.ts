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

export function applyFormula(formula: string, sequence: number, date: Date, padLength: number): string {
  if (!formula || formula.indexOf('{seq') === -1) {
    throw new Error('Formula must include {seq} or {seq:n} so each item gets a distinct code.');
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
  return formula.replace(/\{(yyyy|yy|MM|dd|HH|mm|seq(?::\d+)?)\}/g, (match: string, token: string) => {
    if (token === 'seq') {
      return pad(sequence, padLength);
    }
    if (token.indexOf('seq:') === 0) {
      const width = Number(token.slice(4));
      if (!Number.isInteger(width) || width < 1 || width > 12) {
        throw new Error('Sequence width must be from 1 to 12.');
      }
      return pad(sequence, width);
    }
    return values[token] || match;
  });
}

export function planNextCode(input: {
  formula: string;
  resetPeriod: ResetPeriod | string;
  lastResetDate?: string;
  currentCount: number;
  padLength: number;
  now: Date;
}): { code: string; sequence: number; lastResetDate: string } {
  const rolled = periodRolledOver(input.resetPeriod, input.lastResetDate, input.now);
  const sequence = (rolled ? 0 : input.currentCount) + 1;
  return {
    code: applyFormula(input.formula, sequence, input.now, input.padLength),
    sequence,
    lastResetDate: input.now.toISOString()
  };
}

export function isBlankNumber(value: unknown): boolean {
  if (value === null || value === undefined) {
    return true;
  }
  return String(value).trim().length === 0;
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
