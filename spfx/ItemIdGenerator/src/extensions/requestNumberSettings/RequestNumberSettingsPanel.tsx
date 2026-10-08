import * as React from 'react';
import {
  Checkbox,
  DefaultButton,
  Dropdown,
  IDropdownOption,
  MessageBar,
  MessageBarType,
  Panel,
  PanelType,
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
  ensureConfigList,
  ensureTargetNumberColumn,
  loadConfigRows,
  saveConfigRow
} from '../../webparts/itemIdGenerator/services/configListClient';
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
  draft?: IConfigDraft;
  errorMessage: string;
  warningMessage: string;
  statusMessage: string;
}

export default class RequestNumberSettingsPanel
  extends React.Component<IRequestNumberSettingsPanelProps, IRequestNumberSettingsPanelState> {

  public constructor(props: IRequestNumberSettingsPanelProps) {
    super(props);
    this.state = {
      loading: true,
      saving: false,
      canEdit: false,
      errorMessage: '',
      warningMessage: '',
      statusMessage: ''
    };
  }

  public componentDidMount(): void {
    this._load().catch((error: unknown) => {
      this.setState({ loading: false, errorMessage: messageOf(error) });
    });
  }

  public render(): React.ReactElement {
    return (
      <Panel
        isOpen={true}
        type={PanelType.medium}
        headerText={`Request numbers for ${this.props.listTitle}`}
        onDismiss={this.props.onDismiss}
        isBlocking={false}
      >
        <p className={styles.intro}>
          Set the formula and save. Add an item on this list. The next number is written into the number column. Refresh the list if it is still blank.
        </p>
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
        {!this.state.loading && this.state.draft && this._renderForm(this.state.draft)}
      </Panel>
    );
  }

  private _renderForm(draft: IConfigDraft): React.ReactElement {
    const busy = this.state.saving || !this.state.canEdit;
    const padText = Number.isInteger(draft.padLength) ? String(draft.padLength) : '';
    return (
      <div className={styles.panel}>
        {!this.state.canEdit && (
          <div className={styles.status}>
            <MessageBar messageBarType={MessageBarType.info}>
              You can view these settings. Edit permission on the config list is required to change them.
            </MessageBar>
          </div>
        )}
        <TextField label="Title" required={true} value={draft.title} onChange={this._onTitleChange} disabled={busy} />
        <TextField label="This list" value={this.props.listUrl} readOnly={true} />
        <TextField label="List GUID" value={draft.targetListGuid} readOnly={true} />
        <TextField
          label="Number column internal name"
          required={true}
          description="Single-line text column on this list. Example: RequestNumber."
          value={draft.numberColumnInternalName}
          onChange={this._onFieldChange}
          disabled={busy}
        />
        <TextField
          label="Formula"
          required={true}
          description="UTC tokens: {yyyy} {yy} {MM} {dd} {HH} {mm} {seq} {seq:n}."
          value={draft.formula}
          onChange={this._onFormulaChange}
          disabled={busy}
        />
        <TextField
          label="Pad length"
          value={padText}
          onChange={this._onPadChange}
          disabled={busy}
          description="Width for {seq}. Use 0 through 12."
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
          <TextField label="Current count" value={String(draft.currentCount || 0)} disabled={true} />
        )}
        {this.state.canEdit && (
          <div className={styles.actions}>
            <PrimaryButton text={this.state.saving ? 'Saving...' : 'Save settings'} onClick={this._onSave} disabled={this.state.saving} />
            <DefaultButton text="Close" onClick={this.props.onDismiss} disabled={this.state.saving} />
          </div>
        )}
      </div>
    );
  }

  private _onTitleChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this._patch({ title: value || '' });
  };

  private _onFieldChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this._patch({ numberColumnInternalName: value || '' });
  };

  private _onFormulaChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this._patch({ formula: value || '' });
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
    const rows = await loadConfigRows(this.props.spHttpClient, this.props.siteAbsoluteUrl, listTitle);
    const existing = findConfigForList(rows, this.props.listGuid);
    const draft = existing || draftForCurrentList(this.props.listTitle, this.props.listUrl, this.props.listGuid);
    let canEdit = false;
    try {
      canEdit = await canEditConfigList(this.props.spHttpClient, this.props.siteAbsoluteUrl, listTitle);
    } catch {
      canEdit = false;
    }
    this.setState({
      loading: false,
      draft,
      canEdit
    });
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
      ? 'Saved. Add an item, then refresh this list. The number appears in the number column.'
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
        if (written > 0) {
          statusMessage = `Saved. ${written} item${written === 1 ? '' : 's'} received a number. Refresh the list to see them.`;
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
  return error instanceof Error ? error.message : 'The request number settings could not be saved.';
}
