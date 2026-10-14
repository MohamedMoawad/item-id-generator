import { useState } from 'react';
import { IChildRow, IPendingAttachment, IRepeatingRowsConfig, IValidationIssue } from '../models/repeatingRows';
import { changeCell, createRow, duplicateRow, moveRow, withSort } from '../utilities/rows';
import { validateRows } from '../utilities/validation';

export interface IRepeatingRowsState {
  rows: IChildRow[];
  issues: IValidationIssue[];
  attachments: IPendingAttachment[];
  add: () => void;
  edit: (key: string, internalName: string, value: string) => void;
  remove: (key: string) => void;
  duplicate: (key: string) => void;
  move: (fromIndex: number, toIndex: number) => void;
  attach: (rowKey: string, fileName: string, content: ArrayBuffer) => void;
  append: (page: IChildRow[]) => void;
  check: () => IValidationIssue[];
}

export function useRepeatingRows(config: IRepeatingRowsConfig, initial: IChildRow[], parentId: number): IRepeatingRowsState {
  const [rows, setRows] = useState<IChildRow[]>(initial);
  const [issues, setIssues] = useState<IValidationIssue[]>([]);
  const [attachments, setAttachments] = useState<IPendingAttachment[]>([]);

  const add = (): void => {
    if (config.maxRows > 0 && rows.length >= config.maxRows) {
      return;
    }
    setRows(withSort([...rows, createRow(config, parentId, rows.length + 1)]));
  };

  const edit = (key: string, internalName: string, value: string): void => {
    setRows(rows.map((row) => row.key === key ? changeCell(row, config.columns, internalName, value) : row));
  };

  const remove = (key: string): void => {
    if (rows.length <= config.minRows) {
      return;
    }
    setRows(withSort(rows.filter((row) => row.key !== key)));
    setAttachments(attachments.filter((file) => file.rowKey !== key));
  };

  const duplicate = (key: string): void => {
    if (config.maxRows > 0 && rows.length >= config.maxRows) {
      return;
    }
    const index = rows.findIndex((row) => row.key === key);
    setRows(duplicateRow(rows, index));
  };

  const move = (fromIndex: number, toIndex: number): void => {
    setRows(moveRow(rows, fromIndex, toIndex));
  };

  const attach = (rowKey: string, fileName: string, content: ArrayBuffer): void => {
    setAttachments([...attachments, { rowKey, fileName, content }]);
    setRows(rows.map((row) => {
      if (row.key !== rowKey) {
        return row;
      }
      const names = config.columns.filter((column) => column.type === 'attachment');
      if (names.length === 0) {
        return row;
      }
      const current = row.values[names[0].internalName] || '';
      const next = current ? `${current}, ${fileName}` : fileName;
      return changeCell(row, config.columns, names[0].internalName, next);
    }));
  };

  const append = (page: IChildRow[]): void => {
    const seen = new Set(rows.map((row) => row.id));
    const extra = page.filter((row) => row.id === undefined || !seen.has(row.id));
    setRows(withSort([...rows, ...extra]));
  };

  const check = (): IValidationIssue[] => {
    const next = validateRows(rows, config, attachments);
    setIssues(next);
    return next;
  };

  return { rows, issues, attachments, add, edit, remove, duplicate, move, attach, append, check };
}
