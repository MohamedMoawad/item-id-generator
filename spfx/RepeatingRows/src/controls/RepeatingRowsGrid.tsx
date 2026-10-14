import * as React from 'react';
import { Button } from '@fluentui/react-components';
import { IChildRow, ILookupOption, IPersonOption, IRepeatingRowsConfig, IValidationIssue } from '../models/repeatingRows';
import { columnTotal } from '../utilities/formulas';
import { columnIsVisible } from '../utilities/rows';
import { visibleRange } from '../utilities/batch';
import { IRowEditorProps, RowEditor } from './RowEditor';
import styles from './repeatingRows.module.scss';

const ROW_HEIGHT = 72;

export interface IRepeatingRowsGridProps {
  config: IRepeatingRowsConfig;
  rows: IChildRow[];
  issues: IValidationIssue[];
  lookups: { [internalName: string]: ILookupOption[] };
  readOnly: boolean;
  hasMore: boolean;
  onAdd: () => void;
  onEdit: (key: string, internalName: string, value: string) => void;
  onDelete: (key: string) => void;
  onDuplicate: (key: string) => void;
  onMove: (fromIndex: number, toIndex: number) => void;
  onAttach: (rowKey: string, file: File) => void;
  onLoadMore: () => void;
  onSearchPeople: (text: string) => Promise<IPersonOption[]>;
  onChoosePerson: (key: string, internalName: string, person: IPersonOption) => void;
}

export function RepeatingRowsGrid(props: IRepeatingRowsGridProps): React.ReactElement {
  const [dragIndex, setDragIndex] = React.useState<number>(-1);
  const [scrollTop, setScrollTop] = React.useState(0);
  const virtual = props.rows.length > 40;
  const range = virtual ? visibleRange(props.rows.length, scrollTop, 480, ROW_HEIGHT) : { start: 0, end: props.rows.length };
  const visibleColumns = props.config.columns.filter((column) => props.rows.some((row) => columnIsVisible(column, row.values)) || !column.visibleWhen);
  const atMax = props.config.maxRows > 0 && props.rows.length >= props.config.maxRows;
  return (
    <section>
      <div className={styles.toolbar}>
        <Button appearance="primary" disabled={props.readOnly || atMax} onClick={props.onAdd}>Add row</Button>
        {props.hasMore && <Button appearance="secondary" onClick={props.onLoadMore}>Load more</Button>}
      </div>
      <div
        className={`${styles.gridWrap} ${virtual ? styles.virtual : ''}`}
        onScroll={(event) => setScrollTop(event.currentTarget.scrollTop)}
      >
        {virtual && <div style={{ height: range.start * ROW_HEIGHT }} />}
        <table className={styles.grid}>
          <thead>
            <tr>
              <th aria-label="Reorder" />
              {visibleColumns.map((column) => <th key={column.internalName}>{column.title}</th>)}
              <th>Actions</th>
            </tr>
          </thead>
          <tbody>
            {props.rows.slice(range.start, range.end).map((row, offset) => {
              const index = range.start + offset;
              const editor: IRowEditorProps = {
                row,
                columns: visibleColumns,
                lookups: props.lookups,
                issues: props.issues.filter((issue) => issue.rowKey === row.key),
                readOnly: props.readOnly || (!props.config.allowEdit && row.id !== undefined),
                allowDelete: props.config.allowDelete && props.rows.length > props.config.minRows,
                allowDuplicate: !atMax,
                onChange: (internalName, value) => props.onEdit(row.key, internalName, value),
                onDelete: () => props.onDelete(row.key),
                onDuplicate: () => props.onDuplicate(row.key),
                onSearchPeople: props.onSearchPeople,
                onChoosePerson: (internalName, person) => props.onChoosePerson(row.key, internalName, person),
                onAttach: (file) => props.onAttach(row.key, file)
              };
              return (
                <tr
                  key={row.key}
                  draggable={!props.readOnly}
                  onDragStart={() => setDragIndex(index)}
                  onDragOver={(event) => event.preventDefault()}
                  onDrop={() => {
                    props.onMove(dragIndex, index);
                    setDragIndex(-1);
                  }}
                >
                  <td className={styles.handle} aria-label="Drag to reorder">::</td>
                  <RowEditor {...editor} />
                </tr>
              );
            })}
          </tbody>
        </table>
        {virtual && <div style={{ height: Math.max(0, props.rows.length - range.end) * ROW_HEIGHT }} />}
      </div>
      <div className={styles.totals}>
        {props.config.columns.filter((column) => column.formula || column.type === 'currency' || column.type === 'number').map((column) => (
          <span key={column.internalName}>{column.title} total: {formatTotal(columnTotal(props.rows, column.internalName))}</span>
        ))}
      </div>
    </section>
  );
}

function formatTotal(value: number): string {
  return String(Math.round(value * 100) / 100);
}
