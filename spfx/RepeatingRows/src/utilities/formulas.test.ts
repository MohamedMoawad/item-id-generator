import { columnTotal, evaluateFormula } from './formulas';

describe('formulas', () => {
  it('multiplies quantity by unit price', () => {
    expect(evaluateFormula('{Quantity}*{UnitPrice}', { Quantity: '2', UnitPrice: '100' })).toBe('200');
  });

  it('rejects anything that is not arithmetic', () => {
    expect(evaluateFormula('{Quantity};alert(1)', { Quantity: '2' })).toBe('');
  });

  it('totals a column', () => {
    const rows = [{ values: { Amount: '200' } }, { values: { Amount: '50' } }];
    expect(columnTotal(rows, 'Amount')).toBe(250);
  });
});
