import { IChildRow, IRepeatingColumn, IRepeatingRowsConfig, IValidationIssue } from '../models/repeatingRows';
import { columnIsVisible } from './rows';

export function validateRows(
  rows: IChildRow[],
  config: IRepeatingRowsConfig,
  attachments?: { rowKey: string }[]
): IValidationIssue[] {
  const issues: IValidationIssue[] = [];
  if (rows.length < config.minRows) {
    issues.push({ rowKey: '', message: `Add at least ${config.minRows} row${config.minRows === 1 ? '' : 's'}.` });
  }
  if (config.maxRows > 0 && rows.length > config.maxRows) {
    issues.push({ rowKey: '', message: `This column allows at most ${config.maxRows} rows.` });
  }
  rows.forEach((row) => {
    config.columns.forEach((column) => {
      if (!columnIsVisible(column, row.values) || column.formula) {
        return;
      }
      const value = (row.values[column.internalName] || '').trim();
      if (column.type === 'attachment') {
        const pending = (attachments || []).some((file) => file.rowKey === row.key);
        if (column.required && !value && !pending) {
          issues.push({ rowKey: row.key, column: column.internalName, message: `${column.title} is required.` });
        }
        return;
      }
      if (column.required && !value) {
        issues.push({ rowKey: row.key, column: column.internalName, message: `${column.title} is required.` });
        return;
      }
      if (value && !valueMatches(column, value)) {
        issues.push({ rowKey: row.key, column: column.internalName, message: `${column.title} is not valid.` });
      }
    });
  });
  return issues;
}

function valueMatches(column: IRepeatingColumn, value: string): boolean {
  if (column.type === 'number' || column.type === 'currency') {
    return Number.isFinite(Number(value));
  }
  if (column.type === 'choice' && column.choices && column.choices.length > 0) {
    return column.choices.indexOf(value) >= 0;
  }
  if (column.type === 'date') {
    return /^\d{4}-\d{2}-\d{2}$/.test(value);
  }
  if (column.type === 'boolean') {
    return value === 'true' || value === 'false' || value === '1' || value === '0';
  }
  return true;
}
