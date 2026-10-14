import { defaultRepeatingConfig } from '../models/repeatingRows';
import { changeCell, columnIsVisible, duplicateRow, moveRow, resetRowKeys, seedRows } from './rows';

describe('rows', () => {
  beforeEach(() => resetRowKeys());

  it('seeds the minimum number of blank rows', () => {
    const config = defaultRepeatingConfig('Details');
    config.minRows = 2;
    const rows = seedRows(config, 0);
    expect(rows).toHaveLength(2);
    expect(rows[0].values.Quantity).toBe('1');
    expect(rows[0].values.Amount).toBe('0');
  });

  it('reorders and duplicates without keeping the source id', () => {
    const config = defaultRepeatingConfig('Details');
    const rows = seedRows(config, 4);
    rows[0].id = 9;
    const duplicated = duplicateRow(rows, 0);
    expect(duplicated[1].id).toBeUndefined();
    expect(duplicated[1].values.Description).toBe(rows[0].values.Description);
    const moved = moveRow(duplicated, 0, 1);
    expect(moved[1].id).toBe(9);
    expect(moved.map((row) => row.sortOrder)).toEqual([1, 2]);
  });

  it('recalculates a formula when a cell changes', () => {
    const config = defaultRepeatingConfig('Details');
    const row = seedRows(config, 1)[0];
    const next = changeCell(row, config.columns, 'Quantity', '2');
    const priced = changeCell(next, config.columns, 'UnitPrice', '100');
    expect(priced.values.Amount).toBe('200');
  });

  it('hides a column until the condition matches', () => {
    const column = { internalName: 'Note', title: 'Note', type: 'text' as const, required: false, visibleWhen: 'Status=Open' };
    expect(columnIsVisible(column, { Status: 'Closed' })).toBe(false);
    expect(columnIsVisible(column, { Status: 'Open' })).toBe(true);
  });
});
