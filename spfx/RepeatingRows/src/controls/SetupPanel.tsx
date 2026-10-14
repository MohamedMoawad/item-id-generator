import * as React from 'react';
import { Button, Checkbox, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Field, Input, MessageBar, MessageBarBody, Spinner, Textarea } from '@fluentui/react-components';
import { defaultRepeatingConfig, IRepeatingRowsConfig } from '../models/repeatingRows';
import { parseRepeatingConfig } from '../utilities/configParser';
import { toUserMessage } from '../utilities/errors';
import styles from './repeatingRows.module.scss';

export interface ISetupPanelProps {
  listTitle: string;
  onDismiss: () => void;
  onProvision: (config: IRepeatingRowsConfig) => Promise<void>;
}

export function SetupPanel(props: ISetupPanelProps): React.ReactElement {
  const starter = defaultRepeatingConfig(`${props.listTitle} Details`);
  const [childListTitle, setChildListTitle] = React.useState(starter.childListTitle);
  const [relationshipField, setRelationshipField] = React.useState(starter.relationshipField);
  const [minRows, setMinRows] = React.useState(String(starter.minRows));
  const [maxRows, setMaxRows] = React.useState(String(starter.maxRows));
  const [allowDelete, setAllowDelete] = React.useState(starter.allowDelete);
  const [allowEdit, setAllowEdit] = React.useState(starter.allowEdit);
  const [columnsText, setColumnsText] = React.useState(JSON.stringify(starter.columns, undefined, 2));
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
        relationshipKind: 'lookup',
        minRows: Number(minRows),
        maxRows: Number(maxRows),
        allowDelete,
        allowEdit,
        columns: JSON.parse(columnsText) as unknown
      }, starter.childListTitle);
    } catch (error) {
      setBusy(false);
      setMessage(toUserMessage(error));
      return;
    }
    props.onProvision(config).then(() => {
      setBusy(false);
      setMessage('Repeating rows are ready on this list. New and edit forms now open the line-item grid.');
    }).catch((error: unknown) => {
      setBusy(false);
      setMessage(toUserMessage(error));
    });
  };

  return (
    <Dialog open={true} onOpenChange={(_event, data) => { if (!data.open) { props.onDismiss(); } }}>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>Repeating Rows</DialogTitle>
          <DialogContent className={styles.setup}>
            <p>
              SharePoint Online does not let an app add a new entry to the built-in column type picker.
              This command creates a Line Items column, a child list, and replaces the new, edit, and display forms.
            </p>
            <Field label="Child list name" required>
              <Input value={childListTitle} onChange={(_event, data) => setChildListTitle(data.value)} />
            </Field>
            <Field label="Relationship field" required>
              <Input value={relationshipField} onChange={(_event, data) => setRelationshipField(data.value)} />
            </Field>
            <Field label="Min rows">
              <Input type="number" value={minRows} onChange={(_event, data) => setMinRows(data.value)} />
            </Field>
            <Field label="Max rows (0 means no limit)">
              <Input type="number" value={maxRows} onChange={(_event, data) => setMaxRows(data.value)} />
            </Field>
            <Checkbox checked={allowDelete} onChange={(_event, data) => setAllowDelete(!!data.checked)} label="Allow row delete" />
            <Checkbox checked={allowEdit} onChange={(_event, data) => setAllowEdit(!!data.checked)} label="Allow row edit" />
            <Field label="Columns" hint="JSON array. Include required, choices, lookupListTitle, formula, cascadeFrom, and visibleWhen.">
              <Textarea value={columnsText} rows={12} onChange={(_event, data) => setColumnsText(data.value)} />
            </Field>
            {message && (
              <MessageBar className={styles.banner}>
                <MessageBarBody>{message}</MessageBarBody>
              </MessageBar>
            )}
            {busy && <Spinner label="Creating the child list and form" />}
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={props.onDismiss}>Close</Button>
            <Button appearance="primary" disabled={busy} onClick={save}>Create column</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
