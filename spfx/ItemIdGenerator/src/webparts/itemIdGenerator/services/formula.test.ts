import {
  applyFormula,
  columnMapForItem,
  columnNamesInFormula,
  columnToken,
  formulaColumns,
  itemQueryParts,
  periodKeyFor,
  periodRolledOver,
  planNextCode
} from './formula';

const october = new Date('2026-10-06T15:04:00Z');

describe('formula', () => {
  it('pads the sequence and formats UTC tokens', () => {
    expect(applyFormula('REQ-{yyyy}{MM}-{seq:4}', 12, october, 0)).toBe('REQ-202610-0012');
    expect(applyFormula('REQ-{seq}', 12, october, 4)).toBe('REQ-0012');
    expect(applyFormula('REQ-{seq}', 12, october, 0)).toBe('REQ-12');
    expect(applyFormula('REQ-{seq:2}', 12, october, 6)).toBe('REQ-12');
    expect(applyFormula('{yyyy}-{department}-{seq:2}', 3, october, 0)).toBe('2026-{department}-03');
  });

  it('rejects a code that would repeat', () => {
    expect(() => applyFormula('REQ-{yyyy}', 1, october, 0)).toThrow(/\{counter\}/);
  });

  it('uses counter and list column values, including a lookup title', () => {
    expect(applyFormula('REQ-{counter}', 12, october, 4)).toBe('REQ-0012');
    expect(applyFormula('{Title}-{Department.title}-{counter:2}', 3, october, 4, {
      Title: 'Desk',
      'Department.title': 'IT'
    })).toBe('Desk-IT-03');
    expect(applyFormula('{Department}-{counter}', 1, october, 0, {
      Department: 'IT'
    })).toBe('IT-1');
  });

  it('uses UTC period keys', () => {
    expect(periodKeyFor('Day', october)).toBe('2026-10-06');
    expect(periodKeyFor('month', october)).toBe('2026-10');
    expect(periodKeyFor('YEAR', october)).toBe('2026');
    expect(periodKeyFor('None', october)).toBe('');
    expect(periodRolledOver('None', '2020-01-01T00:00:00.000Z', october)).toBe(false);
    expect(periodRolledOver('Month', '', october)).toBe(false);
    expect(periodRolledOver('Month', 'not-a-date', october)).toBe(false);
    expect(periodRolledOver('Month', '2026-09-30T23:00:00.000Z', october)).toBe(true);
    expect(periodRolledOver('Month', '2026-10-01T00:00:00.000Z', october)).toBe(false);
  });

  it('plans the next code from the current count', () => {
    const next = planNextCode({
      formula: 'REQ-{seq}',
      resetPeriod: 'Month',
      lastResetDate: '2026-10-01T00:00:00.000Z',
      currentCount: 4,
      padLength: 4,
      now: october
    });
    expect(next).toEqual({
      code: 'REQ-0005',
      sequence: 5,
      lastResetDate: october.toISOString()
    });
  });

  it('builds a lookup token and reads the lookup title from the item', () => {
    const columns = formulaColumns([
      { InternalName: 'Title', Title: 'Title', TypeAsString: 'Text', Hidden: false },
      { InternalName: 'Department', Title: 'Department', TypeAsString: 'Lookup', Hidden: false },
      { InternalName: 'LinkTitle', Title: 'Title', TypeAsString: 'Computed', Hidden: false }
    ]);
    expect(columnToken('Department', 'Lookup')).toBe('{Department.title}');
    expect(columns.map((column) => column.token)).toEqual(['{Department.title}', '{Title}']);
    expect(columnNamesInFormula('REQ-{Department.title}-{counter}')).toEqual(['Department']);
    expect(itemQueryParts('RequestNumber', ['Department', 'Title'], columns)).toEqual({
      select: ['Id', 'RequestNumber', 'Department/Title', 'Department/Id', 'Title'],
      expand: ['Department']
    });
    expect(columnMapForItem({
      Title: 'Desk',
      Department: { Title: 'IT', Id: 4 }
    }, ['Title', 'Department'], columns)).toMatchObject({
      Title: 'Desk',
      'Department.title': 'IT',
      'Department.Id': '4'
    });
  });
});
