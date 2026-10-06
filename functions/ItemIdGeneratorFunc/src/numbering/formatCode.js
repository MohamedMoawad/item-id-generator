'use strict';

const TOKEN = /\{(yyyy|yy|MM|dd|HH|mm|seq(?::\d+)?)\}/g;

function pad(value, width) {
  return String(value).padStart(width, '0');
}

function normalizeResetPeriod(value) {
  const text = String(value ?? '').trim().toLowerCase();
  if (text === 'day' || text === 'month' || text === 'year') {
    return text;
  }
  return 'none';
}

function periodKeyFor(resetPeriod, date) {
  const period = normalizeResetPeriod(resetPeriod);
  const year = date.getUTCFullYear();
  const month = pad(date.getUTCMonth() + 1, 2);
  const day = pad(date.getUTCDate(), 2);
  if (period === 'day') {
    return `${year}-${month}-${day}`;
  }
  if (period === 'month') {
    return `${year}-${month}`;
  }
  if (period === 'year') {
    return String(year);
  }
  return '';
}

function applyFormula(formula, input) {
  if (typeof formula !== 'string' || formula.indexOf('{seq') === -1) {
    throw new Error('Formula must include {seq} or {seq:n} so each item gets a distinct code.');
  }
  const date = input.date;
  const sequence = input.sequence;
  const values = {
    yyyy: String(date.getUTCFullYear()),
    yy: String(date.getUTCFullYear()).slice(-2),
    MM: pad(date.getUTCMonth() + 1, 2),
    dd: pad(date.getUTCDate(), 2),
    HH: pad(date.getUTCHours(), 2),
    mm: pad(date.getUTCMinutes(), 2)
  };

  return formula.replace(TOKEN, (match, token) => {
    if (token === 'seq') {
      const width = Number(input.padLength) || 0;
      if (width < 0 || width > 12) {
        throw new Error('PadLength must be from 0 to 12.');
      }
      return width > 0 ? pad(sequence, width) : String(sequence);
    }
    if (token.indexOf('seq:') === 0) {
      const width = Number(token.slice(4));
      if (!Number.isInteger(width) || width < 1 || width > 12) {
        throw new Error('Sequence width must be from 1 to 12.');
      }
      return pad(sequence, width);
    }
    return Object.prototype.hasOwnProperty.call(values, token) ? values[token] : match;
  });
}

function periodRolledOver(resetPeriod, lastResetDate, now) {
  const period = normalizeResetPeriod(resetPeriod);
  if (period === 'none' || !lastResetDate) {
    return false;
  }
  const last = new Date(lastResetDate);
  if (Number.isNaN(last.getTime())) {
    return false;
  }
  return periodKeyFor(period, last) !== periodKeyFor(period, now);
}

module.exports = {
  applyFormula,
  normalizeResetPeriod,
  periodKeyFor,
  periodRolledOver
};
