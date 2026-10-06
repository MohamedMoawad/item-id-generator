'use strict';

const assert = require('node:assert/strict');
const { describe, it } = require('node:test');
const { applyFormula, periodKeyFor, periodRolledOver } = require('./formatCode');

const october = new Date('2026-10-06T15:04:00Z');

describe('applyFormula', () => {
  it('pads the sequence and formats UTC date tokens', () => {
    const code = applyFormula('REQ-{yyyy}{MM}-{seq:4}', { sequence: 12, date: october });
    assert.equal(code, 'REQ-202610-0012');
  });

  it('uses PadLength for {seq} and lets {seq:n} override it', () => {
    assert.equal(applyFormula('REQ-{seq}', { sequence: 12, date: october, padLength: 4 }), 'REQ-0012');
    assert.equal(applyFormula('REQ-{seq}', { sequence: 12, date: october, padLength: 0 }), 'REQ-12');
    assert.equal(applyFormula('REQ-{seq:2}', { sequence: 12, date: october, padLength: 6 }), 'REQ-12');
  });

  it('leaves tokens it does not understand in place', () => {
    const code = applyFormula('{yyyy}-{department}-{seq:2}', { sequence: 3, date: october });
    assert.equal(code, '2026-{department}-03');
  });

  it('rejects a formula that would mint the same code for every item', () => {
    assert.throws(
      () => applyFormula('REQ-{yyyy}', { sequence: 1, date: october }),
      /\{seq\}/
    );
  });
});

describe('periodKeyFor', () => {
  it('uses UTC day, month, and year keys', () => {
    assert.equal(periodKeyFor('Day', october), '2026-10-06');
    assert.equal(periodKeyFor('month', october), '2026-10');
    assert.equal(periodKeyFor('YEAR', october), '2026');
    assert.equal(periodKeyFor('None', october), '');
  });
});

describe('periodRolledOver', () => {
  it('resets only when a previous reset date falls in a different UTC period', () => {
    assert.equal(periodRolledOver('None', '2020-01-01T00:00:00.000Z', october), false);
    assert.equal(periodRolledOver('Month', '', october), false);
    assert.equal(periodRolledOver('Month', 'not-a-date', october), false);
    assert.equal(periodRolledOver('Month', '2026-09-30T23:00:00.000Z', october), true);
    assert.equal(periodRolledOver('Month', '2026-10-01T00:00:00.000Z', october), false);
  });
});
