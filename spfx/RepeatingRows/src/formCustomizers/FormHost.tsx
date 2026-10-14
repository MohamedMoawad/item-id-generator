import * as React from 'react';
import { Button, FluentProvider, MessageBar, MessageBarBody, Spinner, webLightTheme } from '@fluentui/react-components';
import { FormDisplayMode } from '@microsoft/sp-core-library';
import { IChildRow, ILookupOption, IParentField, IParentValues, IPendingAttachment, IPersonOption, IRepeatingRowsConfig } from '../models/repeatingRows';
import { ParentFields } from '../controls/ParentFields';
import { RepeatingRowsGrid } from '../controls/RepeatingRowsGrid';
import { useRepeatingRows } from '../hooks/useRepeatingRows';
import { seedRows } from '../utilities/rows';
import { validateParent } from '../utilities/parentValues';
import { toUserMessage } from '../utilities/errors';
import styles from '../controls/repeatingRows.module.scss';

export interface IFormHostProps {
  displayMode: FormDisplayMode;
  config: IRepeatingRowsConfig;
  fields: IParentField[];
  initialValues: IParentValues;
  initialRows: IChildRow[];
  etag?: string;
  parentId?: number;
  lookups: { [internalName: string]: ILookupOption[] };
  hasMore: boolean;
  onLoadMore: (afterSortOrder: number) => Promise<{ rows: IChildRow[]; hasMore: boolean }>;
  onSearchPeople: (text: string) => Promise<IPersonOption[]>;
  onResolvePerson: (person: IPersonOption) => Promise<string>;
  onSave: (values: IParentValues, rows: IChildRow[], attachments: IPendingAttachment[], etag?: string) => Promise<void>;
  onClose: () => void;
}

export function FormHost(props: IFormHostProps): React.ReactElement {
  const readOnly = props.displayMode === FormDisplayMode.Display;
  const seeded = props.initialRows.length > 0 || readOnly ? props.initialRows : seedRows(props.config, props.parentId || 0);
  const grid = useRepeatingRows(props.config, seeded, props.parentId || 0);
  const [values, setValues] = React.useState<IParentValues>(props.initialValues);
  const [more, setMore] = React.useState(props.hasMore);
  const [message, setMessage] = React.useState('');
  const [busy, setBusy] = React.useState(false);
  const title = props.displayMode === FormDisplayMode.New ? 'New item' : (readOnly ? 'Item' : 'Edit item');

  const save = (): void => {
    const parentIssue = validateParent(props.fields, values);
    const issues = grid.check();
    if (parentIssue || issues.length > 0) {
      setMessage(parentIssue || issues[0].message);
      return;
    }
    setBusy(true);
    setMessage('');
    props.onSave(values, grid.rows, grid.attachments, props.etag).then(() => {
      setBusy(false);
    }).catch((error: unknown) => {
      setBusy(false);
      setMessage(toUserMessage(error));
    });
  };

  return (
    <FluentProvider theme={webLightTheme} className={styles.root}>
      <h2>{title}</h2>
      {message && (
        <MessageBar className={styles.banner} intent="error">
          <MessageBarBody>{message}</MessageBarBody>
        </MessageBar>
      )}
      <ParentFields
        fields={props.fields}
        values={values}
        readOnly={readOnly}
        onChange={(internalName, value) => setValues({ ...values, [internalName]: value })}
        onSearchPeople={props.onSearchPeople}
        onChoosePerson={(internalName, person) => {
          props.onResolvePerson(person).then((stored) => {
            setValues({ ...values, [internalName]: stored });
          }).catch((error: unknown) => setMessage(toUserMessage(error)));
        }}
      />
      <RepeatingRowsGrid
        config={props.config}
        rows={grid.rows}
        issues={grid.issues}
        lookups={props.lookups}
        readOnly={readOnly}
        hasMore={more}
        onAdd={grid.add}
        onEdit={grid.edit}
        onDelete={grid.remove}
        onDuplicate={grid.duplicate}
        onMove={grid.move}
        onAttach={(rowKey, file) => {
          readFile(file).then((content) => grid.attach(rowKey, file.name, content)).catch((error: unknown) => {
            setMessage(toUserMessage(error));
          });
        }}
        onLoadMore={() => {
          const last = grid.rows.length > 0 ? grid.rows[grid.rows.length - 1].sortOrder : 0;
          props.onLoadMore(last).then((page) => {
            grid.append(page.rows);
            setMore(page.hasMore);
          }).catch((error: unknown) => setMessage(toUserMessage(error)));
        }}
        onSearchPeople={props.onSearchPeople}
        onChoosePerson={(key, internalName, person) => {
          props.onResolvePerson(person).then((stored) => grid.edit(key, internalName, stored)).catch((error: unknown) => {
            setMessage(toUserMessage(error));
          });
        }}
      />
      <div className={styles.toolbar}>
        {!readOnly && <Button appearance="primary" disabled={busy} onClick={save}>{busy ? 'Saving' : 'Save'}</Button>}
        <Button appearance="secondary" onClick={props.onClose}>Close</Button>
        {busy && <Spinner size="tiny" label="Saving rows" />}
      </div>
    </FluentProvider>
  );
}

function readFile(file: File): Promise<ArrayBuffer> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as ArrayBuffer);
    reader.onerror = () => reject(new Error('The file could not be read.'));
    reader.readAsArrayBuffer(file);
  });
}
