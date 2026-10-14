import { IChildRow } from '../models/repeatingRows';
import { planChildChanges } from './savePlan';

function row(id: number | undefined, description: string, sortOrder: number): IChildRow {
  return { key: `k${id || 'new'}`, id, parentId: 3, sortOrder, values: { Description: description } };
}

describe('save plan', () => {
  it('creates new rows, updates edits and reorders, and deletes missing ids', () => {
    const original = [row(1, 'Item 1', 1), row(2, 'Item 2', 2)];
    const current = [row(2, 'Item 2', 1), row(undefined, 'Item 3', 2)];
    const plan = planChildChanges(original, current);
    expect(plan.deleteIds).toEqual([1]);
    expect(plan.create).toHaveLength(1);
    expect(plan.update.map((item) => item.id)).toEqual([2]);
  });
});
