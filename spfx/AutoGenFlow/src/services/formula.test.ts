import {
  applyFormula,
  columnTokensInFormula,
  padForFlow,
  padSequence,
  periodRolledOver,
  planNextCode,
  prepareRule
} from './formula';

const october = new Date('2026-10-16T15:04:00Z');

describe('formula', () => {
  it('pads with the same digits the flow expression uses', () => {
    expect(padForFlow(1, 4)).toBe('0001');
    expect(padForFlow(22, 4)).toBe('0022');
    expect(padSequence(1, 5)).toBe('00001');
    expect(padForFlow(10000, 4)).toBe('10000');
    expect(padForFlow(7, 0)).toBe('7');
    for (let width = 0; width <= 12; width += 1) {
      expect(padForFlow(1, width)).toBe(padSequence(1, width));
      expect(padForFlow(11, width)).toBe(padSequence(11, width));
    }
  });

  it('writes the counter once', () => {
    const prepared = prepareRule('{counter:4}{counter}', 4);
    expect(prepared.formula).toBe('{counter:4}');
    expect(prepared.counterToken).toBe('{counter:4}');
    expect(prepared.counterWidth).toBe(4);
    expect(applyFormula(prepared.formula, 1, october, 4)).toBe('0001');
    expect(applyFormula('{counter:4}{Id}', 2, october, 4)).toBe('0002');
    expect(planNextCode({
      formula: '{counter:5}',
      resetPeriod: 'None',
      currentCount: '1',
      padLength: 0,
      now: october
    }).code).toBe('00002');
  });

  it('keeps column tokens for the flow and resets on a new month', () => {
    const prepared = prepareRule('REQ-{yyyy}{MM}-{Title}-{Department.title}-{counter:4}', 4);
    expect(prepared.formulaColumns).toBe('Title,Department.title');
    expect(columnTokensInFormula(prepared.formula)).toEqual(['Title', 'Department.title']);
    expect(planNextCode({
      formula: prepared.formula,
      resetPeriod: 'Month',
      lastResetDate: '2026-09-30T00:00:00Z',
      currentCount: 40,
      padLength: 4,
      now: october,
      columns: { Title: 'Desk', 'Department.title': 'IT' }
    })).toMatchObject({ sequence: 1, code: 'REQ-202610-Desk-IT-0001' });
    expect(periodRolledOver('None', '2020-01-01T00:00:00Z', october)).toBe(false);
  });
});
