import * as React from 'react';
import { Button, Checkbox, Field, FluentProvider, Input, MessageBar, MessageBarBody, Spinner, webLightTheme } from '@fluentui/react-components';
import { defaultRepeatingConfig, IRepeatingColumn, IRepeatingRowsConfig, RelationshipKind } from '../models/repeatingRows';
import { blankColumn, cleanColumn } from '../utilities/columnDraft';
import { parseRepeatingConfig } from '../utilities/configParser';
import { toUserMessage } from '../utilities/errors';
import { ColumnDesigner } from './ColumnDesigner';
import styles from './repeatingRows.module.scss';

export interface ISetupPanelProps {
  listTitle: string;
  ready: boolean;
  existing?: IRepeatingRowsConfig;
  onDismiss: () => void;
  onProvision: (config: IRepeatingRowsConfig) => Promise<void>;
}

export function SetupPanel(props: ISetupPanelProps): React.ReactElement {
  const dialogRef = React.useRef<HTMLDivElement | null>(null);
  React.useEffect(() => {
    const onKey = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') {
        props.onDismiss();
      }
    };
    document.addEventListener('keydown', onKey);
    if (dialogRef.current) {
      dialogRef.current.focus();
    }
    return (): void => {
      document.removeEventListener('keydown', onKey);
    };
  }, [props]);

  return (
    <div
      className={styles.overlay}
      onClick={(event) => {
        if (event.target === event.currentTarget) {
          props.onDismiss();
        }
      }}
    >
      <div
        ref={dialogRef}
        className={styles.dialog}
        role="dialog"
        aria-modal="true"
        aria-label="Repeating Rows"
        tabIndex={-1}
      >
        <FluentProvider theme={webLightTheme} className={styles.dialogBody}>
          <h2 className={styles.dialogTitle}>Repeating Rows</h2>
          {props.ready ? <SetupForm {...props} /> : <Spinner label="Reading the current column" />}
        </FluentProvider>
      </div>
    </div>
  );
}

function SetupForm(props: ISetupPanelProps): React.ReactElement {
  const starter = props.existing || defaultRepeatingConfig(`${props.listTitle} Details`);
  const [childListTitle, setChildListTitle] = React.useState(starter.childListTitle);
  const [relationshipField, setRelationshipField] = React.useState(starter.relationshipField);
  const [relationshipKind, setRelationshipKind] = React.useState<RelationshipKind>(starter.relationshipKind);
  const [minRows, setMinRows] = React.useState(String(starter.minRows));
  const [maxRows, setMaxRows] = React.useState(String(starter.maxRows));
  const [allowDelete, setAllowDelete] = React.useState(starter.allowDelete);
  const [allowEdit, setAllowEdit] = React.useState(starter.allowEdit);
  const [columns, setColumns] = React.useState<IRepeatingColumn[]>(starter.columns.map((column) => ({ ...column })));
  const [message, setMessage] = React.useState('');
  const [busy, setBusy] = React.useState(false);

  const save = (): void => {
    setBusy(true);
    setMessage('');
    let config: IRepeatingRowsConfig;
    try {
      config = parseRepeatingConfig({
        childListTitle,
        relationshipField,
        relationshipKind,
        minRows: Number(minRows),
        maxRows: Number(maxRows),
        allowDelete,
        allowEdit,
        columns: columns.map(cleanColumn)
      }, starter.childListTitle);
    } catch (error) {
      setBusy(false);
      setMessage(toUserMessage(error));
      return;
    }
    props.onProvision(config).then(() => {
      setBusy(false);
      setMessage(props.existing
        ? 'The column settings are saved. New and edit forms use this child list.'
        : 'Repeating rows are ready on this list. New and edit forms now open the line-item grid.');
    }).catch((error: unknown) => {
      setBusy(false);
      setMessage(toUserMessage(error));
    });
  };

  return (
    <div className={styles.setup}>
      <p>
        SharePoint Online does not add this to the built-in column type picker.
        This command creates the Line Items column, the child list, and the new, edit, and display forms.
      </p>
      <Field label="Child list name" required>
        <Input value={childListTitle} onChange={(_event, data) => setChildListTitle(data.value)} />
      </Field>
      <Field label="Relationship field" required>
        <Input value={relationshipField} onChange={(_event, data) => setRelationshipField(data.value)} />
      </Field>
      <Field label="Relationship">
        <select aria-label="Relationship" value={relationshipKind} onChange={(event) => setRelationshipKind(event.target.value as RelationshipKind)}>
          <option value="lookup">Lookup with cascade delete</option>
          <option value="number">Number</option>
        </select>
      </Field>
      <Field label="Min rows">
        <Input type="number" value={minRows} onChange={(_event, data) => setMinRows(data.value)} />
      </Field>
      <Field label="Max rows (0 means no limit)">
        <Input type="number" value={maxRows} onChange={(_event, data) => setMaxRows(data.value)} />
      </Field>
      <Checkbox checked={allowDelete} onChange={(_event, data) => setAllowDelete(!!data.checked)} label="Allow row delete" />
      <Checkbox checked={allowEdit} onChange={(_event, data) => setAllowEdit(!!data.checked)} label="Allow row edit" />
      <div className={styles.columnHead}>
        <strong>Columns to display</strong>
        <Button appearance="secondary" onClick={() => setColumns([...columns, blankColumn(columns.length + 1)])}>Add column</Button>
      </div>
      <ColumnDesigner columns={columns} onChange={setColumns} />
      {message && (
        <MessageBar className={styles.banner}>
          <MessageBarBody>{message}</MessageBarBody>
        </MessageBar>
      )}
      {busy && <Spinner label="Saving the child list and form" />}
      <div className={styles.dialogActions}>
        <Button appearance="secondary" onClick={props.onDismiss}>Close</Button>
        <Button appearance="primary" disabled={busy} onClick={save}>{props.existing ? 'Save column' : 'Create column'}</Button>
      </div>
    </div>
  );
}
