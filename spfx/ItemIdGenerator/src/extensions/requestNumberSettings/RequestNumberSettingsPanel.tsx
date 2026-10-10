import * as React from 'react';
import {
  Checkbox,
  DefaultButton,
  Dialog,
  DialogType,
  Dropdown,
  IContextualMenuItem,
  IDropdownOption,
  IconButton,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  Spinner,
  SpinnerSize,
  TextField
} from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import styles from './RequestNumberSettingsPanel.module.scss';
import {
  DEFAULT_CONFIG_LIST_TITLE,
  deriveWebFromListUrl,
  draftForCurrentList,
  findConfigForList,
  IConfigDraft,
  RESET_PERIODS,
  ResetPeriod,
  validateDraft
} from '../../webparts/itemIdGenerator/services/configContract';
import {
  canEditConfigList,
  deleteConfigRow,
  ensureConfigList,
  ensureTargetNumberColumn,
  hideConfigList,
  loadConfigRows,
  loadFormulaColumns,
  saveConfigRow
} from '../../webparts/itemIdGenerator/services/configListClient';
import { IFormulaColumn } from '../../webparts/itemIdGenerator/services/formula';
import { numberBlankItems } from '../../webparts/itemIdGenerator/services/listNumbering';

export interface IRequestNumberSettingsPanelProps {
  siteAbsoluteUrl: string;
  configListTitle: string;
  listTitle: string;
  listUrl: string;
  listGuid: string;
  spHttpClient: SPHttpClient;
  onDismiss: () => void;
}

interface IRequestNumberSettingsPanelState {
  loading: boolean;
  saving: boolean;
  canEdit: boolean;
  editing: boolean;
  draft?: IConfigDraft;
  columns: IFormulaColumn[];
  errorMessage: string;
  warningMessage: string;
  statusMessage: string;
  confirmDelete: boolean;
}

export default class RequestNumberSettingsPanel
  extends React.Component<IRequestNumberSettingsPanelProps, IRequestNumberSettingsPanelState> {

  public constructor(props: IRequestNumberSettingsPanelProps) {
    super(props);
    this.state = {
      loading: true,
      saving: false,
      canEdit: false,
      editing: false,
      columns: [],
      errorMessage: '',
      warningMessage: '',
      statusMessage: '',
      confirmDelete: false
    };
  }

  public componentDidMount(): void {
    this._load().catch((error: unknown) => {
      this.setState({ loading: false, errorMessage: messageOf(error) });
    });
  }

  public render(): React.ReactElement {
    const draft = this.state.draft;
    const showRow = !!draft && draft.id !== undefined && !this.state.editing;
    return (
      <Dialog
        hidden={false}
        onDismiss={this.props.onDismiss}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Autogen Feature',
          showCloseButton: true
        }}
        modalProps={{
          isBlocking: false,
          styles: { main: { maxWidth: 880, width: '92%' } }
        }}
      >
        <div className={styles.toolbar}>
          <p className={styles.statusLine}>{this._statusLine()}</p>
          <DefaultButton
            className={styles.settingsButton}
            text="Settings"
            iconProps={{ iconName: 'Settings' }}
            menuIconProps={{ iconName: 'ChevronDown' }}
            menuProps={{ items: this._settingsItems() }}
            disabled={this.state.loading || !draft}
          />
        </div>
        {this.state.loading && <Spinner size={SpinnerSize.small} label="Loading settings" />}
        {this.state.errorMessage && (
          <div className={styles.status}>
            <MessageBar messageBarType={MessageBarType.error}>{this.state.errorMessage}</MessageBar>
          </div>
        )}
        {this.state.warningMessage && (
          <div className={styles.status}>
            <MessageBar messageBarType={MessageBarType.warning}>{this.state.warningMessage}</MessageBar>
          </div>
        )}
        {this.state.statusMessage && (
          <div className={styles.status}>
            <MessageBar messageBarType={MessageBarType.success}>{this.state.statusMessage}</MessageBar>
          </div>
        )}
        {showRow && draft && this._renderSavedRow(draft)}
        {!this.state.loading && draft && !showRow && this._renderForm(draft)}
      </Dialog>
    );
  }

  private _renderSavedRow(draft: IConfigDraft): React.ReactElement {
    return (
      <div className={styles.table}>
        <div className={styles.head}>
          <span>Title</span>
          <span>Number column</span>
          <span>Formula</span>
          <span>Status</span>
          <span>Owner</span>
          <span />
        </div>
        <div className={styles.dataRow}>
          <span className={styles.titleLink}>{draft.title || this.props.listTitle}</span>
          <span>{this._columnTitle(draft.numberColumnInternalName)}</span>
          <span>{draft.formula}</span>
          <span>{draft.isActive ? 'Active' : 'Off'}</span>
          <span>{draft.modifiedBy || ''}</span>
          <IconButton
            iconProps={{ iconName: 'More' }}
            ariaLabel="Autogen actions"
            menuProps={{ items: this._rowItems() }}
            disabled={this.state.saving}
          />
        </div>
        {this.state.confirmDelete && (
          <div className={styles.confirm}>
            <span>Delete Autogen from this list?</span>
            <PrimaryButton text={this.state.saving ? 'Deleting...' : 'Delete'} onClick={this._onDelete} disabled={this.state.saving} />
            <DefaultButton text="Cancel" onClick={this._onCancelDelete} disabled={this.state.saving} />
          </div>
        )}
      </div>
    );
  }

  private _renderForm(draft: IConfigDraft): React.ReactElement {
    const editing = this.state.editing && this.state.canEdit;
    const busy = this.state.saving || !editing;
    const padText = Number.isInteger(draft.padLength) ? String(draft.padLength) : '';
    return (
      <div className={styles.panel}>
        <p className={styles.intro}>
          Choose an existing single line of text column. Create that column in the list first if it is not listed. {'{counter}'} is the next number.
        </p>
        <TextField label="Last modified" value={formatModified(draft.modified)} readOnly={true} />
        <TextField label="Modified by" value={draft.modifiedBy || 'Not saved yet'} readOnly={true} />
        {!this.state.canEdit && (
          <div className={styles.status}>
            <MessageBar messageBarType={MessageBarType.info}>
              You can view these settings. Permission to edit this list is required to change them.
            </MessageBar>
          </div>
        )}
        {this._textColumns().length === 0 && (
          <div className={styles.status}>
            <MessageBar messageBarType={MessageBarType.warning}>
              This list has no single line of text column. Create one in list settings, then open Autogen Setting again.
            </MessageBar>
          </div>
        )}
        <Dropdown
          label="Number column"
          required={true}
          placeholder="Select a single line of text column"
          selectedKey={draft.numberColumnInternalName || undefined}
          options={this._numberColumnOptions()}
          onChange={this._onNumberColumnChange}
          disabled={busy || this._textColumns().length === 0}
        />
        <TextField
          label="Formula"
          required={true}
          description="Use {counter} or {counter:4}. Dates: {yyyy} {yy} {MM} {dd}. Columns: {Title} or {Department.title}."
          value={draft.formula}
          onChange={this._onFormulaChange}
          disabled={busy}
        />
        <Dropdown
          label="Add a column to the formula"
          placeholder="Insert a column from this list"
          options={this.state.columns
            .filter((column) => column.internalName !== draft.numberColumnInternalName.trim())
            .map((column) => ({ key: column.token, text: `${column.title} ${column.token}` }))}
          onChange={this._onInsertColumn}
          disabled={busy || this.state.columns.length === 0}
        />
        <TextField
          label="Pad length"
          value={padText}
          onChange={this._onPadChange}
          disabled={busy}
          description="Width for {counter}. {counter:4} uses its own width. Use 0 through 12."
        />
        <Dropdown
          label="Reset period"
          selectedKey={draft.resetPeriod}
          options={RESET_PERIODS.map((period) => ({ key: period, text: period }))}
          onChange={this._onResetChange}
          disabled={busy}
        />
        <Checkbox label="Active" checked={draft.isActive} onChange={this._onActiveChange} disabled={busy} />
        {draft.id !== undefined && (
          <TextField label="Current counter" value={String(draft.currentCount || 0)} readOnly={true} />
        )}
        <div className={styles.actions}>
          {editing && (
            <PrimaryButton text={this.state.saving ? 'Saving...' : 'Save settings'} onClick={this._onSave} disabled={this.state.saving} />
          )}
          <DefaultButton text="Close" onClick={this.props.onDismiss} disabled={this.state.saving} />
        </div>
      </div>
    );
  }

  private _statusLine(): string {
    const draft = this.state.draft;
    if (this.state.loading) {
      return 'Loading settings.';
    }
    if (!draft || !draft.modified) {
      return 'Not saved yet.';
    }
    const who = draft.modifiedBy ? ` by ${draft.modifiedBy}` : '';
    return `Last modified ${formatModified(draft.modified)}${who}.`;
  }

  private _settingsItems(): IContextualMenuItem[] {
    const editing = this.state.editing && this.state.canEdit;
    return [
      {
        key: 'edit',
        text: 'Edit',
        iconProps: { iconName: 'Edit' },
        disabled: !this.state.canEdit || editing || this.state.saving,
        onClick: this._onEdit
      },
      {
        key: 'view',
        text: 'View',
        iconProps: { iconName: 'View' },
        disabled: !this.state.editing || this.state.saving,
        onClick: this._onView
      }
    ];
  }

  private _columnTitle(internalName: string): string {
    const match = this.state.columns.filter((column) => column.internalName === internalName)[0];
    return match ? match.title : internalName;
  }

  private _rowItems(): IContextualMenuItem[] {
    const locked = !this.state.canEdit || this.state.saving;
    return [
      {
        key: 'edit',
        text: 'Edit',
        iconProps: { iconName: 'Edit' },
        disabled: locked,
        onClick: this._onEdit
      },
      {
        key: 'delete',
        text: 'Delete',
        iconProps: { iconName: 'Delete' },
        disabled: locked,
        onClick: this._onAskDelete
      }
    ];
  }

  private _textColumns(): IFormulaColumn[] {
    return this.state.columns.filter((column) => column.type === 'Text');
  }

  private _numberColumnOptions(): IDropdownOption[] {
    return this._textColumns().map((column) => ({ key: column.internalName, text: column.title }));
  }

  private _onNumberColumnChange = (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (!option) {
      return;
    }
    this._patch({ numberColumnInternalName: String(option.key) });
  };

  private _onEdit = (): void => {
    this.setState({ editing: true, confirmDelete: false, statusMessage: '', errorMessage: '' });
  };

  private _onAskDelete = (): void => {
    this.setState({ confirmDelete: true, errorMessage: '' });
  };

  private _onCancelDelete = (): void => {
    this.setState({ confirmDelete: false });
  };

  private _onDelete = (): void => {
    this._delete().catch((error: unknown) => {
      this.setState({ saving: false, errorMessage: messageOf(error) });
    });
  };

  private _onView = (): void => {
    this._load().catch((error: unknown) => {
      this.setState({ loading: false, errorMessage: messageOf(error) });
    });
  };

  private _onFormulaChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this._patch({ formula: value || '' });
  };

  private _onInsertColumn = (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (!option || !this.state.draft) {
      return;
    }
    const token = String(option.key);
    const formula = this.state.draft.formula || '';
    this._patch({ formula: formula.indexOf(token) >= 0 ? formula : `${formula}${token}` });
  };

  private _onPadChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    const text = (value || '').trim();
    this._patch({ padLength: /^\d+$/.test(text) ? Number(text) : Number.NaN });
  };

  private _onResetChange = (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    const key = option ? String(option.key) : 'None';
    const resetPeriod: ResetPeriod = key === 'Day' || key === 'Month' || key === 'Year' ? key : 'None';
    this._patch({ resetPeriod });
  };

  private _onActiveChange = (_event?: React.FormEvent<HTMLElement | HTMLInputElement>, checked?: boolean): void => {
    this._patch({ isActive: !!checked });
  };

  private _onSave = (): void => {
    this._save().catch((error: unknown) => {
      this.setState({ saving: false, errorMessage: messageOf(error) });
    });
  };

  private _patch(patch: Partial<IConfigDraft>): void {
    if (!this.state.draft) {
      return;
    }
    this.setState({ draft: { ...this.state.draft, ...patch } });
  }

  private async _load(): Promise<void> {
    const listTitle = this.props.configListTitle.trim() || DEFAULT_CONFIG_LIST_TITLE;
    this.setState({ loading: true, errorMessage: '' });
    await ensureConfigList(this.props.spHttpClient, this.props.siteAbsoluteUrl, listTitle);
    await hideConfigList(this.props.spHttpClient, this.props.siteAbsoluteUrl, listTitle);
    const rows = await loadConfigRows(this.props.spHttpClient, this.props.siteAbsoluteUrl, listTitle);
    const existing = findConfigForList(rows, this.props.listGuid);
    const draft = existing || draftForCurrentList(this.props.listTitle, this.props.listUrl, this.props.listGuid);
    let canEdit = false;
    let columns: IFormulaColumn[] = [];
    try {
      canEdit = await canEditConfigList(this.props.spHttpClient, this.props.siteAbsoluteUrl, listTitle);
    } catch {
      canEdit = false;
    }
    try {
      columns = await loadFormulaColumns(
        this.props.spHttpClient,
        deriveWebFromListUrl(this.props.listUrl).webAbsoluteUrl,
        this.props.listGuid
      );
    } catch {
      columns = [];
    }
    const selectedIsText = columns.some((column) => column.type === 'Text' && column.internalName === draft.numberColumnInternalName);
    if (!selectedIsText) {
      draft.numberColumnInternalName = '';
    }
    this.setState({
      loading: false,
      draft,
      canEdit,
      editing: !existing && canEdit,
      columns,
      confirmDelete: false
    });
  }

  private async _delete(): Promise<void> {
    if (!this.state.draft || this.state.draft.id === undefined || !this.state.canEdit) {
      return;
    }
    const listTitle = this.props.configListTitle.trim() || DEFAULT_CONFIG_LIST_TITLE;
    this.setState({ saving: true, errorMessage: '', statusMessage: '', warningMessage: '' });
    await deleteConfigRow(
      this.props.spHttpClient,
      this.props.siteAbsoluteUrl,
      listTitle,
      this.state.draft.id
    );
    this.setState({ saving: false, statusMessage: 'Removed Autogen from this list.' });
    await this._load();
  }

  private async _save(): Promise<void> {
    if (!this.state.draft || !this.state.canEdit) {
      return;
    }
    const problem = validateDraft(this.state.draft);
    if (problem) {
      this.setState({ errorMessage: problem, statusMessage: '', warningMessage: '' });
      return;
    }
    const listTitle = this.props.configListTitle.trim() || DEFAULT_CONFIG_LIST_TITLE;
    this.setState({ saving: true, errorMessage: '', statusMessage: '', warningMessage: '' });
    await ensureTargetNumberColumn(
      this.props.spHttpClient,
      this.state.draft.targetListUrl,
      this.state.draft.targetListGuid,
      this.state.draft.numberColumnInternalName
    );
    await saveConfigRow(
      this.props.spHttpClient,
      this.props.siteAbsoluteUrl,
      listTitle,
      this.state.draft
    );
    let statusMessage = this.state.draft.isActive
      ? 'Saved. Add an item. The number appears on the list by itself.'
      : 'Saved the settings for this list. Numbering is turned off.';
    let warningMessage = '';
    if (this.state.draft.isActive) {
      try {
        const webUrl = deriveWebFromListUrl(this.state.draft.targetListUrl).webAbsoluteUrl;
        const written = await numberBlankItems(
          this.props.spHttpClient,
          this.props.siteAbsoluteUrl,
          listTitle,
          webUrl,
          this.state.draft.targetListGuid
        );
        if (written.length > 0) {
          statusMessage = `Saved. ${written.length} item${written.length === 1 ? '' : 's'} received a number.`;
        }
      } catch (error) {
        warningMessage = messageOf(error);
      }
    }
    this.setState({
      saving: false,
      statusMessage,
      warningMessage
    });
    await this._load();
  }
}

function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : 'The Autogen settings could not be saved.';
}

function formatModified(value?: string): string {
  if (!value) {
    return 'Not saved yet';
  }
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return value;
  }
  return date.toLocaleString();
}
