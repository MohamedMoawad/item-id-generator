import * as React from 'react';
import { Field, Input, Textarea } from '@fluentui/react-components';
import { IParentField, IParentValues, IPersonOption } from '../models/repeatingRows';
import styles from './repeatingRows.module.scss';

export interface IParentFieldsProps {
  fields: IParentField[];
  values: IParentValues;
  readOnly: boolean;
  onChange: (internalName: string, value: string) => void;
  onSearchPeople: (text: string) => Promise<IPersonOption[]>;
  onChoosePerson: (internalName: string, person: IPersonOption) => void;
}

export function ParentFields(props: IParentFieldsProps): React.ReactElement {
  return (
    <div className={styles.parentGrid}>
      {props.fields.map((field) => (
        <ParentField key={field.internalName} field={field} props={props} />
      ))}
    </div>
  );
}

function ParentField(input: { field: IParentField; props: IParentFieldsProps }): React.ReactElement {
  const { field, props } = input;
  const value = String(props.values[field.internalName] || '');
  const change = (next: string): void => props.onChange(field.internalName, next);
  if (field.type === 'note') {
    return (
      <Field label={field.title} required={field.required}>
        <Textarea aria-label={field.title} value={value} readOnly={props.readOnly} onChange={(_event, data) => change(data.value)} />
      </Field>
    );
  }
  if (field.type === 'choice') {
    return (
      <Field label={field.title} required={field.required}>
        <select aria-label={field.title} value={value} disabled={props.readOnly} onChange={(event) => change(event.target.value)}>
          <option value="">Select</option>
          {(field.choices || []).map((choice) => <option key={choice} value={choice}>{choice}</option>)}
        </select>
      </Field>
    );
  }
  if (field.type === 'boolean') {
    return (
      <Field label={field.title}>
        <input
          aria-label={field.title}
          type="checkbox"
          checked={value === 'true'}
          disabled={props.readOnly}
          onChange={(event) => change(event.target.checked ? 'true' : 'false')}
        />
      </Field>
    );
  }
  if (field.type === 'person') {
    return <PersonField field={field} value={value} props={props} />;
  }
  const inputType = field.type === 'date' ? 'date' : (field.type === 'number' || field.type === 'currency' ? 'number' : 'text');
  return (
    <Field label={field.title} required={field.required}>
      <Input
        aria-label={field.title}
        type={inputType}
        value={value}
        readOnly={props.readOnly}
        onChange={(_event, data) => change(data.value)}
      />
    </Field>
  );
}

function PersonField(input: { field: IParentField; value: string; props: IParentFieldsProps }): React.ReactElement {
  const [people, setPeople] = React.useState<IPersonOption[]>([]);
  const label = input.value.indexOf('|') >= 0 ? input.value.split('|').slice(1).join('|') : input.value;
  return (
    <Field label={input.field.title} required={input.field.required}>
      <Input
        aria-label={input.field.title}
        value={label}
        readOnly={input.props.readOnly}
        onChange={(_event, data) => {
          input.props.onChange(input.field.internalName, data.value);
          input.props.onSearchPeople(data.value).then(setPeople).catch(() => setPeople([]));
        }}
      />
      <span className={styles.suggestions}>
        {people.map((person) => (
          <button
            key={person.login || person.text}
            type="button"
            onClick={() => {
              input.props.onChoosePerson(input.field.internalName, person);
              setPeople([]);
            }}
          >
            {person.text}
          </button>
        ))}
      </span>
    </Field>
  );
}
