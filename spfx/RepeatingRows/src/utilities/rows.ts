import { IChildRow, IRepeatingColumn, IRepeatingRowsConfig } from '../models/repeatingRows';
import { applyFormulas } from './formulas';

let rowSequence = 1;

export function nextRowKey(): string {
  rowSequence += 1;
  return `row-${rowSequence}`;
}

export function blankValues(columns: IRepeatingColumn[]): { [internalName: string]: string } {
  const values: { [internalName: string]: string } = {};
  columns.forEach((column) => {
    values[column.internalName] = column.defaultValue || '';
  });
  return applyFormulas(values, columns);
}

export function createRow(config: IRepeatingRowsConfig, parentId: number, sortOrder: number): IChildRow {
  return {
    key: nextRowKey(),
    parentId,
    sortOrder,
    values: blankValues(config.columns)
  };
}

export function resetRowKeys(): void {
  rowSequence = 1;
}

export function seedRows(config: IRepeatingRowsConfig, parentId: number): IChildRow[] {
  const count = Math.max(config.minRows, 0);
  const rows: IChildRow[] = [];
  for (let index = 0; index < count; index += 1) {
    rows.push(createRow(config, parentId, index + 1));
  }
  return withSort(rows);
}

export function withSort(rows: IChildRow[]): IChildRow[] {
  return rows.map((row, index) => ({ ...row, sortOrder: index + 1 }));
}

export function moveRow(rows: IChildRow[], fromIndex: number, toIndex: number): IChildRow[] {
  if (fromIndex < 0 || toIndex < 0 || fromIndex >= rows.length || toIndex >= rows.length || fromIndex === toIndex) {
    return rows;
  }
  const next = rows.slice();
  const [moved] = next.splice(fromIndex, 1);
  next.splice(toIndex, 0, moved);
  return withSort(next);
}

export function duplicateRow(rows: IChildRow[], index: number): IChildRow[] {
  const source = rows[index];
  if (!source) {
    return rows;
  }
  const copy: IChildRow = {
    key: nextRowKey(),
    parentId: source.parentId,
    sortOrder: source.sortOrder,
    values: { ...source.values }
  };
  const next = rows.slice();
  next.splice(index + 1, 0, copy);
  return withSort(next);
}

export function changeCell(row: IChildRow, columns: IRepeatingColumn[], internalName: string, value: string): IChildRow {
  const values = applyFormulas({ ...row.values, [internalName]: value }, columns);
  return { ...row, values };
}

export function columnIsVisible(column: IRepeatingColumn, values: { [internalName: string]: string }): boolean {
  if (!column.visibleWhen || column.visibleWhen.indexOf('=') < 0) {
    return true;
  }
  const splitAt = column.visibleWhen.indexOf('=');
  const field = column.visibleWhen.slice(0, splitAt).trim();
  const expected = column.visibleWhen.slice(splitAt + 1).trim();
  return (values[field] || '') === expected;
}
