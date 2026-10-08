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
  draftForCurrentList,
  findConfigForList,
  IConfigDraft,
  isPlaceholderSetting,
  RESET_PERIODS,
  ResetPeriod,
  validateDraft
} from '../../webparts/itemIdGenerator/services/configContract';
import {
  callRegisterWebhook,
  canEditConfigList,
  ensureConfigList,
  loadConfigRows,
  readRegisterWebhookUrl,
  saveConfigRow,
  saveRegisterWebhookUrl
} from '../../webparts/itemIdGenerator/services/configListClient';

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
  registerWebhookUrl: string;
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
      registerWebhookUrl: '',
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
    const listTitle = this.props.configListTitle.trim() || DEFAULT_CONFIG_LIST_TITLE;
    return (
      <Panel
        isOpen={true}
        type={PanelType.medium}
        headerText={`Request numbers for ${this.props.listTitle}`}
        onDismiss={this.props.onDismiss}
        isBlocking={false}
      >
        <p className={styles.intro}>
          These settings are stored on {listTitle} for this list only. The same app numbers every site collection where it is added. You do not enter a site name.
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
        <TextField
          label="RegisterWebhook URL"
          value={this.state.registerWebhookUrl}
          onChange={this._onWebhookChange}
          disabled={busy}
          description="Paste this once for the organization. Include the function key. A SharePoint administrator save stores it for every site collection."
        />
        {draft.id !== undefined && (
          <TextField label="Current count" value={String(draft.currentCount || 0)} disabled={true} />
        )}
        {draft.webhookSubscriptionId && (
          <TextField label="Webhook subscription ID" value={draft.webhookSubscriptionId} disabled={true} />
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

  private _onWebhookChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this.setState({ registerWebhookUrl: value || '' });
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
    let registerWebhookUrl = '';
    try {
      canEdit = await canEditConfigList(this.props.spHttpClient, this.props.siteAbsoluteUrl, listTitle);
    } catch {
      canEdit = false;
    }
    try {
      registerWebhookUrl = await readRegisterWebhookUrl(this.props.spHttpClient, this.props.siteAbsoluteUrl);
    } catch {
      registerWebhookUrl = '';
    }
    this.setState({ loading: false, draft, canEdit, registerWebhookUrl });
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
    const itemId = await saveConfigRow(
      this.props.spHttpClient,
      this.props.siteAbsoluteUrl,
      listTitle,
      this.state.draft
    );
    const warnings: string[] = [];
    const webhookUrl = this.state.registerWebhookUrl.trim();
    if (!isPlaceholderSetting(webhookUrl)) {
      try {
        const saved = await saveRegisterWebhookUrl(this.props.spHttpClient, this.props.siteAbsoluteUrl, webhookUrl);
        if (!saved.organization) {
          warnings.push('The URL is saved on this site collection. A SharePoint administrator should save it once so every site collection uses the same RegisterWebhook URL.');
        }
      } catch (error) {
        warnings.push(`Saved the list settings. The RegisterWebhook URL was not stored. ${messageOf(error)}`);
      }
    }
    const registration = await this._registerIfActive(this.state.draft, itemId, webhookUrl);
    if (registration.warningMessage) {
      warnings.push(registration.warningMessage);
    }
    this.setState({
      saving: false,
      statusMessage: registration.statusMessage,
      warningMessage: warnings.join(' ')
    });
    await this._load();
  }

  private async _registerIfActive(
    draft: IConfigDraft,
    itemId: number,
    webhookUrl: string
  ): Promise<{ statusMessage: string; warningMessage: string }> {
    if (!draft.isActive) {
      return { statusMessage: 'Saved the settings for this list. The rule is inactive.', warningMessage: '' };
    }
    if (isPlaceholderSetting(webhookUrl)) {
      return {
        statusMessage: 'Saved the settings for this list.',
        warningMessage: 'Add the RegisterWebhook URL to register the list webhook.'
      };
    }
    try {
      const result = await callRegisterWebhook(webhookUrl, this.props.siteAbsoluteUrl, itemId);
      return {
        statusMessage: result.alreadyRegistered
          ? `Saved. Webhook ${result.subscriptionId} was already registered.`
          : `Saved and registered webhook ${result.subscriptionId}.`,
        warningMessage: ''
      };
    } catch (error) {
      return {
        statusMessage: 'Saved the settings for this list.',
        warningMessage: `Webhook registration failed. ${messageOf(error)}`
      };
    }
  }
}

function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : 'The request number settings could not be saved.';
}
