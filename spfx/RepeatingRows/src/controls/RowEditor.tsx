import * as React from 'react';
import { Button } from '@fluentui/react-components';
import { IChildRow, ILookupOption, IPersonOption, IRepeatingColumn, IValidationIssue } from '../models/repeatingRows';
import { columnIsVisible } from '../utilities/rows';
import styles from './repeatingRows.module.scss';

export interface IRowEditorProps {
  row: IChildRow;
  columns: IRepeatingColumn[];
  lookups: { [internalName: string]: ILookupOption[] };
  issues: IValidationIssue[];
  readOnly: boolean;
  allowDelete: boolean;
  allowDuplicate: boolean;
  onChange: (internalName: string, value: string) => void;
  onDelete: () => void;
  onDuplicate: () => void;
  onSearchPeople: (text: string) => Promise<IPersonOption[]>;
  onChoosePerson: (internalName: string, person: IPersonOption) => void;
  onAttach: (file: File) => void;
}

export function RowEditor(props: IRowEditorProps): React.ReactElement {
  return (
    <>
      {props.columns.filter((column) => columnIsVisible(column, props.row.values)).map((column) => (
        <td key={column.internalName}>
          <Cell column={column} props={props} />
        </td>
      ))}
      <td>
        <div className={styles.actions}>
          {props.allowDuplicate && !props.readOnly && (
            <Button size="small" appearance="subtle" onClick={props.onDuplicate}>Duplicate</Button>
          )}
          {props.allowDelete && !props.readOnly && (
            <Button size="small" appearance="subtle" onClick={props.onDelete}>Delete</Button>
          )}
        </div>
      </td>
    </>
  );
}

function Cell(input: { column: IRepeatingColumn; props: IRowEditorProps }): React.ReactElement {
  const { column, props } = input;
  const value = props.row.values[column.internalName] || '';
  const issue = props.issues.filter((item) => item.column === column.internalName)[0];
  const locked = props.readOnly || column.readOnly === true || !!column.formula;
  return (
    <label className={styles.field}>
      {locked ? <span>{value}</span> : <Editor column={column} value={value} props={props} />}
      {issue && <span className={styles.error}>{issue.message}</span>}
    </label>
  );
}

function Editor(input: { column: IRepeatingColumn; value: string; props: IRowEditorProps }): React.ReactElement {
  const { column, value, props } = input;
  const change = (next: string): void => props.onChange(column.internalName, next);
  if (column.type === 'note' || column.type === 'taxonomy') {
    return (
      <textarea
        aria-label={column.title}
        value={value}
        placeholder={column.type === 'taxonomy' ? 'Label|term-guid' : undefined}
        onChange={(event) => change(event.target.value)}
      />
    );
  }
  if (column.type === 'boolean') {
    return (
      <input
        aria-label={column.title}
        type="checkbox"
        checked={value === 'true' || value === '1'}
        onChange={(event) => change(event.target.checked ? 'true' : 'false')}
      />
    );
  }
  if (column.type === 'choice') {
    return (
      <select aria-label={column.title} value={value} onChange={(event) => change(event.target.value)}>
        <option value="">Select</option>
        {(column.choices || []).map((choice) => <option key={choice} value={choice}>{choice}</option>)}
      </select>
    );
  }
  if (column.type === 'lookup') {
    const parent = column.cascadeFrom ? (props.row.values[column.cascadeFrom] || '') : '';
    const options = (props.lookups[column.internalName] || []).filter((option) => {
      if (!column.cascadeFrom || !option.group) {
        return true;
      }
      return option.group === parent;
    });
    return (
      <select aria-label={column.title} value={value} onChange={(event) => change(event.target.value)}>
        <option value="">Select</option>
        {options.map((option) => <option key={option.id} value={option.id}>{option.text}</option>)}
      </select>
    );
  }
  if (column.type === 'person') {
    return <PersonCell column={column} value={value} props={props} />;
  }
  if (column.type === 'attachment') {
    return (
      <span>
        <span>{value}</span>
        <input
          aria-label={column.title}
          type="file"
          onChange={(event) => {
            const file = event.target.files && event.target.files[0];
            if (file) {
              props.onAttach(file);
            }
          }}
        />
      </span>
    );
  }
  const inputType = column.type === 'date' ? 'date' : (column.type === 'number' || column.type === 'currency' ? 'number' : 'text');
  return (
    <input
      aria-label={column.title}
      type={inputType}
      value={value}
      step={column.type === 'currency' ? '0.01' : undefined}
      onChange={(event) => change(event.target.value)}
    />
  );
}

function PersonCell(input: { column: IRepeatingColumn; value: string; props: IRowEditorProps }): React.ReactElement {
  const [people, setPeople] = React.useState<IPersonOption[]>([]);
  const label = input.value.indexOf('|') >= 0 ? input.value.split('|').slice(1).join('|') : input.value;
  return (
    <span className={styles.field}>
      <input
        aria-label={input.column.title}
        value={label}
        onChange={(event) => {
          const text = event.target.value;
          input.props.onChange(input.column.internalName, text);
          input.props.onSearchPeople(text).then(setPeople).catch(() => setPeople([]));
        }}
      />
      {people.length > 0 && (
        <span className={styles.suggestions}>
          {people.map((person) => (
            <Button
              key={person.login || person.text}
              size="small"
              appearance="subtle"
              onClick={() => {
                input.props.onChoosePerson(input.column.internalName, person);
                setPeople([]);
              }}
            >
              {person.text}
            </Button>
          ))}
        </span>
      )}
    </span>
  );
}
