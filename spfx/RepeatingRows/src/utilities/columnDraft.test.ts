import { purchaseLineColumns } from '../models/repeatingRows';
import { blankColumn, cleanColumn, moveColumn, withChoices } from './columnDraft';

describe('column drafts', () => {
  it('moves a column and drops blank optional values', () => {
    const columns = purchaseLineColumns();
    const moved = moveColumn(columns, 0, 2);
    expect(moved[2].internalName).toBe('Description');
    const cleaned = cleanColumn({
      internalName: ' Status ',
      title: ' Status ',
      type: 'choice',
      required: true,
      defaultValue: '  ',
      formula: ' {Quantity}*2 ',
      choices: [' Open ', '']
    });
    expect(cleaned.internalName).toBe('Status');
    expect(cleaned.defaultValue).toBeUndefined();
    expect(cleaned.formula).toBe('{Quantity}*2');
    expect(cleaned.readOnly).toBe(true);
    expect(cleaned.choices).toEqual(['Open']);
  });

  it('adds a text column and stores comma-separated choices', () => {
    const column = withChoices(blankColumn(3), 'Red, Blue');
    expect(column.internalName).toBe('Column3');
    expect(column.choices).toEqual(['Red', 'Blue']);
  });
});
