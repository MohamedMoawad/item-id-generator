import { IChildRow, ISavePlan } from '../models/repeatingRows';
import { withSort } from './rows';

export function planChildChanges(original: IChildRow[], current: IChildRow[]): ISavePlan {
  const ordered = withSort(current);
  const originalById = new Map<number, IChildRow>();
  original.forEach((row) => {
    if (row.id !== undefined) {
      originalById.set(row.id, row);
    }
  });
  const seen = new Set<number>();
  const create: IChildRow[] = [];
  const update: IChildRow[] = [];
  ordered.forEach((row) => {
    if (row.id === undefined) {
      create.push(row);
      return;
    }
    seen.add(row.id);
    const previous = originalById.get(row.id);
    if (!previous || rowChanged(previous, row)) {
      update.push(row);
    }
  });
  const deleteIds: number[] = [];
  originalById.forEach((_row, id) => {
    if (!seen.has(id)) {
      deleteIds.push(id);
    }
  });
  return { create, update, deleteIds };
}

function rowChanged(previous: IChildRow, next: IChildRow): boolean {
  if (previous.sortOrder !== next.sortOrder) {
    return true;
  }
  const names = Object.keys(next.values);
  for (let index = 0; index < names.length; index += 1) {
    const name = names[index];
    if ((previous.values[name] || '') !== (next.values[name] || '')) {
      return true;
    }
  }
  return false;
}
