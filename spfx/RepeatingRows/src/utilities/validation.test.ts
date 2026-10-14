import { defaultRepeatingConfig, IChildRow } from '../models/repeatingRows';
import { validateRows } from './validation';

function row(values: { [name: string]: string }): IChildRow {
  return { key: 'row-1', parentId: 1, sortOrder: 1, values };
}

describe('validation', () => {
  it('requires the description and a numeric quantity', () => {
    const config = defaultRepeatingConfig('Details');
    const issues = validateRows([row({ Description: '', Quantity: 'two', UnitPrice: '1', Amount: '0' })], config);
    expect(issues.map((issue) => issue.column)).toEqual(['Description', 'Quantity']);
  });

  it('blocks a row count above the maximum', () => {
    const config = defaultRepeatingConfig('Details');
    config.maxRows = 1;
    const issues = validateRows([
      row({ Description: 'A', Quantity: '1', UnitPrice: '1' }),
      { ...row({ Description: 'B', Quantity: '1', UnitPrice: '1' }), key: 'row-2' }
    ], config);
    expect(issues[0].message).toContain('at most 1');
  });
});
