import * as React from 'react';
import { Button, Checkbox, Field, Input } from '@fluentui/react-components';
import { IRepeatingColumn, RepeatingColumnType } from '../models/repeatingRows';
import { choicesText, COLUMN_TYPES, withChoices } from '../utilities/columnDraft';
import styles from './repeatingRows.module.scss';

export interface IColumnDesignerProps {
  columns: IRepeatingColumn[];
  onChange: (columns: IRepeatingColumn[]) => void;
}

export function ColumnDesigner(props: IColumnDesignerProps): React.ReactElement {
  const update = (index: number, column: IRepeatingColumn): void => {
    const next = props.columns.slice();
    next[index] = column;
    props.onChange(next);
  };
  return (
    <div className={styles.columnList}>
      {props.columns.map((column, index) => (
        <section className={styles.columnCard} key={`${column.internalName}-${index}`}>
          <div className={styles.columnHead}>
            <strong>{column.title || 'Column'}</strong>
            <span className={styles.actions}>
              <Button size="small" appearance="subtle" disabled={index === 0} onClick={() => props.onChange(swap(props.columns, index, index - 1))}>Up</Button>
              <Button size="small" appearance="subtle" disabled={index === props.columns.length - 1} onClick={() => props.onChange(swap(props.columns, index, index + 1))}>Down</Button>
              <Button size="small" appearance="subtle" disabled={props.columns.length === 1} onClick={() => props.onChange(props.columns.filter((_item, itemIndex) => itemIndex !== index))}>Remove</Button>
            </span>
          </div>
          <Field label="Display name" required>
            <Input value={column.title} onChange={(_event, data) => update(index, { ...column, title: data.value })} />
          </Field>
          <Field label="Internal name" required>
            <Input value={column.internalName} onChange={(_event, data) => update(index, { ...column, internalName: data.value })} />
          </Field>
          <Field label="Type">
            <select
              aria-label="Type"
              value={column.type}
              onChange={(event) => update(index, { ...column, type: event.target.value as RepeatingColumnType })}
            >
              {COLUMN_TYPES.map((option) => <option key={option.value} value={option.value}>{option.label}</option>)}
            </select>
          </Field>
          <Checkbox
            checked={column.required}
            label="Required"
            onChange={(_event, data) => update(index, { ...column, required: !!data.checked })}
          />
          <Field label="Default value">
            <Input value={column.defaultValue || ''} onChange={(_event, data) => update(index, { ...column, defaultValue: data.value })} />
          </Field>
          {column.type === 'choice' && (
            <Field label="Choices" hint="Separate choices with commas.">
              <Input value={choicesText(column)} onChange={(_event, data) => update(index, withChoices(column, data.value))} />
            </Field>
          )}
          {column.type === 'lookup' && (
            <>
              <Field label="Lookup list" required>
                <Input value={column.lookupListTitle || ''} onChange={(_event, data) => update(index, { ...column, lookupListTitle: data.value })} />
              </Field>
              <Field label="Lookup text field">
                <Input value={column.lookupTextField || ''} placeholder="Title" onChange={(_event, data) => update(index, { ...column, lookupTextField: data.value })} />
              </Field>
              <Field label="Cascade from column">
                <Input value={column.cascadeFrom || ''} placeholder="Category" onChange={(_event, data) => update(index, { ...column, cascadeFrom: data.value })} />
              </Field>
              <Field label="Lookup group field">
                <Input value={column.lookupGroupField || ''} onChange={(_event, data) => update(index, { ...column, lookupGroupField: data.value })} />
              </Field>
            </>
          )}
          <Field label="Formula" hint="Example: {Quantity}*{UnitPrice}">
            <Input value={column.formula || ''} onChange={(_event, data) => update(index, { ...column, formula: data.value })} />
          </Field>
          <Field label="Show when" hint="Example: Status=Open">
            <Input value={column.visibleWhen || ''} onChange={(_event, data) => update(index, { ...column, visibleWhen: data.value })} />
          </Field>
        </section>
      ))}
    </div>
  );
}

function swap(columns: IRepeatingColumn[], fromIndex: number, toIndex: number): IRepeatingColumn[] {
  const next = columns.slice();
  const moved = next[fromIndex];
  next[fromIndex] = next[toIndex];
  next[toIndex] = moved;
  return next;
}
