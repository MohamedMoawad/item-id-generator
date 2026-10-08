import * as React from 'react';
import {
  Checkbox,
  DefaultButton,
  Dropdown,
  IDropdownOption,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  Spinner,
  SpinnerSize,
  TextField
} from '@fluentui/react';
import styles from './ItemIdGenerator.module.scss';
import type { IItemIdGeneratorProps } from './IItemIdGeneratorProps';
import {
  callRegisterWebhook,
  readRegisterWebhookUrl,
  ConfigListRequestError,
  ensureConfigList,
  loadConfigRows,
  resolveTargetList,
  saveConfigRow
} from '../services/configListClient';
import {
  DEFAULT_CONFIG_LIST_TITLE,
  emptyDraft,
  IConfigDraft,
  isPlaceholderSetting,
  RESET_PERIODS,
  ResetPeriod,
  validateDraft
} from '../services/configContract';

interface IItemIdGeneratorState {
  loading: boolean;
  saving: boolean;
  resolving: boolean;
  ensuring: boolean;
  listMissing: boolean;
  items: IConfigDraft[];
  draft?: IConfigDraft;
  errorMessage: string;
  statusMessage: string;
  warningMessage: string;
}

export default class ItemIdGenerator extends React.Component<IItemIdGeneratorProps, IItemIdGeneratorState> {
  public constructor(props: IItemIdGeneratorProps) {
    super(props);
    this.state = {
      loading: true,
      saving: false,
      resolving: false,
      ensuring: false,
      listMissing: false,
      items: [],
      errorMessage: '',
      statusMessage: '',
      warningMessage: ''
    };
  }

  public componentDidMount(): void {
    this._load().catch((error: unknown) => {
      this.setState({ loading: false, errorMessage: messageOf(error) });
    });
  }

  public render(): React.ReactElement<IItemIdGeneratorProps> {
    const listTitle = listTitleOf(this.props);
    const busy = this.state.loading || this.state.saving || this.state.resolving || this.state.ensuring;

    return (
      <section className={styles.itemIdGenerator}>
        <h2 className={styles.header}>AutoGen Feature</h2>
        <p className={styles.lede}>
          {listTitle} lives on this site collection. Add one row per list. Saving an active row
          registers a SharePoint webhook. After that, a new item from the form, the grid, or
          Power Automate is numbered by the Azure Function. This page does not assign numbers.
        </p>
        <div className={styles.toolbar}>
          <PrimaryButton text="New rule" onClick={this._startCreate} disabled={busy} />
          <DefaultButton text={`Ensure ${listTitle}`} onClick={this._onEnsure} disabled={busy} />
          <DefaultButton text="Refresh" onClick={this._onRefresh} disabled={busy} />
        </div>
        {this.state.listMissing && (
          <div className={styles.status}>
            <MessageBar messageBarType={MessageBarType.warning}>
              {listTitle} is not on this site collection yet. Choose Ensure {listTitle} before you add a rule.
            </MessageBar>
          </div>
        )}
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
        {this.state.loading && <Spinner size={SpinnerSize.small} label={`Loading ${listTitle}`} />}
        {!this.state.loading && this.state.items.length === 0 && !this.state.draft && !this.state.listMissing && (
          <p className={styles.lede}>No numbering rules yet. Add one row for each list that should receive request numbers.</p>
        )}
        {!this.state.loading && this.state.items.length > 0 && (
          <div className={styles.cards}>
            {this.state.items.map((item) => (
              <article className={styles.card} key={item.id}>
                <h3 className={styles.cardTitle}>{item.title || 'Untitled rule'}</h3>
                <p className={styles.cardMeta}>
                  {item.isActive ? 'Active' : 'Inactive'}
                  {' · '}
                  Reset {item.resetPeriod}
                  {' · '}
                  Count {item.currentCount || 0}
                  {' · '}
                  Pad {item.padLength}
                </p>
                <p className={styles.cardMeta}>{item.targetListUrl || 'List URL not set'}</p>
                <p className={styles.formula}>{item.targetListGuid || 'GUID not resolved'}</p>
                <p className={styles.cardMeta}>Column {item.numberColumnInternalName || 'not set'}</p>
                <p className={styles.formula}>{item.formula}</p>
                <p className={styles.cardMeta}>
                  Webhook {item.webhookSubscriptionId || 'not registered'}
                </p>
                <div className={styles.actions}>
                  <DefaultButton
                    text="Edit"
                    data-config-id={String(item.id)}
                    onClick={this._onEditClick}
                    disabled={this.state.saving}
                  />
                </div>
              </article>
            ))}
          </div>
        )}
        {this.state.draft && this._renderForm(this.state.draft, busy)}
      </section>
    );
  }

  private _renderForm(draft: IConfigDraft, busy: boolean): React.ReactElement {
    const editing = draft.id !== undefined;
    const padText = Number.isInteger(draft.padLength) ? String(draft.padLength) : '';
    return (
      <div className={styles.form}>
        <TextField label="Title" required={true} value={draft.title} onChange={this._onTitleChange} disabled={busy} />
        <TextField
          label="Target list URL"
          required={true}
          value={draft.targetListUrl}
          onChange={this._onListUrlChange}
          disabled={busy}
          placeholder="https://<tenant>.sharepoint.com/sites/operations/Lists/Requests"
          description="Paste the list or library address, then resolve it. AllItems.aspx is fine."
        />
        <div className={styles.actions}>
          <DefaultButton text={this.state.resolving ? 'Resolving...' : 'Resolve list'} onClick={this._onResolve} disabled={busy} />
        </div>
        <TextField
          label="Target list GUID"
          required={true}
          value={draft.targetListGuid}
          readOnly={true}
          description={draft.targetListTitle ? `Resolved list: ${draft.targetListTitle}` : 'Filled in when you resolve the list URL.'}
        />
        <TextField
          label="Number column internal name"
          required={true}
          description="Single-line text column on the target list. Example: RequestNumber. Create that column yourself."
          value={draft.numberColumnInternalName}
          onChange={this._onFieldChange}
          disabled={busy}
        />
        <TextField
          label="Formula"
          required={true}
          description="UTC tokens: {yyyy} {yy} {MM} {dd} {HH} {mm} {seq} {seq:n}. Must include {seq}."
          value={draft.formula}
          onChange={this._onFormulaChange}
          disabled={busy}
        />
        <TextField
          label="Pad length"
          required={true}
          description="Width for {seq}. {seq:n} overrides this. Use 0 through 12."
          value={padText}
          onChange={this._onPadChange}
          disabled={busy}
        />
        <Dropdown
          label="Reset period"
          selectedKey={draft.resetPeriod}
          options={RESET_PERIODS.map((period) => ({ key: period, text: period }))}
          onChange={this._onResetChange}
          disabled={busy}
        />
        <Checkbox label="Active" checked={draft.isActive} onChange={this._onActiveChange} disabled={busy} />
        {editing && (
          <TextField
            label="Current count"
            value={String(draft.currentCount || 0)}
            disabled={true}
            description="This count increases each time a blank item on the list receives a number."
          />
        )}
        {editing && (
          <TextField label="Webhook subscription ID" value={draft.webhookSubscriptionId || ''} disabled={true} />
        )}
        <div className={styles.actions}>
          <PrimaryButton text={this.state.saving ? 'Saving...' : 'Save rule'} onClick={this._onSave} disabled={busy} />
          <DefaultButton text="Cancel" onClick={this._cancel} disabled={busy} />
        </div>
      </div>
    );
  }

  private _startCreate = (): void => {
    this.setState({ draft: emptyDraft(), errorMessage: '', statusMessage: '', warningMessage: '' });
  };

  private _startEdit = (item: IConfigDraft): void => {
    this.setState({ draft: { ...item }, errorMessage: '', statusMessage: '', warningMessage: '' });
  };

  private _cancel = (): void => {
    this.setState({ draft: undefined, errorMessage: '' });
  };

  private _onRefresh = (): void => {
    this._load().catch((error: unknown) => {
      this.setState({ loading: false, errorMessage: messageOf(error) });
    });
  };

  private _onEnsure = (): void => {
    this._ensure().catch((error: unknown) => {
      this.setState({ ensuring: false, errorMessage: messageOf(error) });
    });
  };

  private _onResolve = (): void => {
    this._resolve().catch((error: unknown) => {
      this.setState({ resolving: false, errorMessage: messageOf(error) });
    });
  };

  private _onSave = (): void => {
    this._save().catch((error: unknown) => {
      this.setState({ saving: false, errorMessage: messageOf(error) });
    });
  };

  private _onEditClick = (event: React.MouseEvent<HTMLElement>): void => {
    const raw = event.currentTarget.getAttribute('data-config-id');
    const id = raw ? Number(raw) : Number.NaN;
    const item = this.state.items.filter((row) => row.id === id)[0];
    if (item) {
      this._startEdit(item);
    }
  };

  private _onTitleChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this._patch({ title: value || '' });
  };

  private _onListUrlChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this._patch({ targetListUrl: value || '', targetListGuid: '', targetListTitle: '' });
  };

  private _onFieldChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this._patch({ numberColumnInternalName: value || '' });
  };

  private _onFormulaChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this._patch({ formula: value || '' });
  };

  private _onPadChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    const text = (value || '').trim();
    if (!/^\d+$/.test(text)) {
      this._patch({ padLength: Number.NaN });
      return;
    }
    this._patch({ padLength: Number(text) });
  };

  private _onResetChange = (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    const key = option ? String(option.key) : 'None';
    const resetPeriod: ResetPeriod = key === 'Day' || key === 'Month' || key === 'Year' ? key : 'None';
    this._patch({ resetPeriod });
  };

  private _onActiveChange = (_event?: React.FormEvent<HTMLElement | HTMLInputElement>, checked?: boolean): void => {
    this._patch({ isActive: !!checked });
  };

  private _patch(patch: Partial<IConfigDraft>): void {
    if (!this.state.draft) {
      return;
    }
    this.setState({ draft: { ...this.state.draft, ...patch } });
  }

  private async _load(): Promise<void> {
    this.setState({ loading: true, errorMessage: '', listMissing: false });
    try {
      const items = await loadConfigRows(
        this.props.spHttpClient,
        this.props.siteAbsoluteUrl,
        listTitleOf(this.props)
      );
      this.setState({ loading: false, items, listMissing: false });
    } catch (error) {
      if (error instanceof ConfigListRequestError && error.status === 404) {
        this.setState({ loading: false, items: [], listMissing: true, errorMessage: '' });
        return;
      }
      throw error;
    }
  }

  private async _ensure(): Promise<void> {
    const listTitle = listTitleOf(this.props);
    this.setState({ ensuring: true, errorMessage: '', statusMessage: '', warningMessage: '' });
    const result = await ensureConfigList(this.props.spHttpClient, this.props.siteAbsoluteUrl, listTitle);
    const created = result.created ? `Created ${listTitle}. ` : `${listTitle} already exists. `;
    const fields = result.fieldsAdded.length > 0
      ? `Added columns: ${result.fieldsAdded.join(', ')}.`
      : 'All columns are already present.';
    this.setState({ ensuring: false, statusMessage: `${created}${fields}` });
    await this._load();
  }

  private async _resolve(): Promise<void> {
    if (!this.state.draft) {
      return;
    }
    this.setState({ resolving: true, errorMessage: '', warningMessage: '' });
    const resolved = await resolveTargetList(this.props.spHttpClient, this.state.draft.targetListUrl);
    const title = this.state.draft.title.trim() ? this.state.draft.title : resolved.title;
    this.setState({
      resolving: false,
      statusMessage: `Resolved ${resolved.title}.`,
      draft: {
        ...this.state.draft,
        title,
        targetListGuid: resolved.id,
        targetListTitle: resolved.title
      }
    });
  }

  private async _save(): Promise<void> {
    if (!this.state.draft) {
      return;
    }
    const problem = validateDraft(this.state.draft);
    if (problem) {
      this.setState({ errorMessage: problem, statusMessage: '', warningMessage: '' });
      return;
    }
    const listTitle = listTitleOf(this.props);
    const draft = this.state.draft;
    this.setState({ saving: true, errorMessage: '', statusMessage: '', warningMessage: '' });
    await ensureConfigList(this.props.spHttpClient, this.props.siteAbsoluteUrl, listTitle);
    const itemId = await saveConfigRow(
      this.props.spHttpClient,
      this.props.siteAbsoluteUrl,
      listTitle,
      draft
    );
    const registration = await this._registerIfActive(draft, itemId);
    this.setState({
      saving: false,
      draft: undefined,
      listMissing: false,
      statusMessage: registration.statusMessage,
      warningMessage: registration.warningMessage
    });
    await this._load();
  }

  private async _registerIfActive(draft: IConfigDraft, itemId: number): Promise<{ statusMessage: string; warningMessage: string }> {
    if (!draft.isActive) {
      return {
        statusMessage: 'Saved the rule. Numbering is turned off.',
        warningMessage: ''
      };
    }
    let webhookUrl = this.props.registerWebhookUrl.trim();
    if (isPlaceholderSetting(webhookUrl)) {
      try {
        webhookUrl = (await readRegisterWebhookUrl(this.props.spHttpClient, this.props.siteAbsoluteUrl)).trim();
      } catch {
        webhookUrl = '';
      }
    }
    if (!isPlaceholderSetting(webhookUrl)) {
      try {
        await callRegisterWebhook(webhookUrl, this.props.siteAbsoluteUrl, itemId);
      } catch {
        // The list still numbers items while it is open.
      }
    }
    return {
      statusMessage: 'Saved the rule. Open the list and add an item. Refresh the list to see the number.',
      warningMessage: ''
    };
  }
}

function listTitleOf(props: IItemIdGeneratorProps): string {
  return props.configListTitle.trim() || DEFAULT_CONFIG_LIST_TITLE;
}

function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : 'The config list request failed.';
}
