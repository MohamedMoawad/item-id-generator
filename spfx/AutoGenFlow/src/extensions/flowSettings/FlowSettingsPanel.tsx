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
import { AadHttpClientFactory, SPHttpClient } from '@microsoft/sp-http';
import {
  DEFAULT_CONFIG_LIST_TITLE,
  draftForCurrentList,
  findConfigForList,
  findSiteFlow,
  IConfigDraft,
  ResetPeriod,
  validateDraft,
  wholeCount
} from '../../services/contract';
import {
  canEditConfigList,
  deleteConfigRow,
  ensureConfigList,
  ensureTargetNumberColumn,
  loadConfigRows,
  loadFormulaColumns,
  saveConfigRow,
  saveFlowBinding,
  showConfigList
} from '../../services/configListClient';
import { chooseConnection, FlowConsentError, flowPortalUrl, IFlowConnection, IFlowEnvironment } from '../../services/flowModel';
import { listConnections, listEnvironments, readFlowProblem, removeListFlow, setRemoteFlowState, syncListFlow } from '../../services/flowClient';
import { formulaInserts, IFormulaColumn, numberColumnChoices, planNextCode, prepareRule } from '../../services/formula';
import { buildItemQuery } from '../../services/itemQuery';
import { IListFlowTarget } from '../../services/siteFlow';
import styles from './FlowSettingsPanel.module.scss';

export interface IFlowSettingsPanelProps {
  siteAbsoluteUrl: string;
  siteTitle: string;
  webAbsoluteUrl: string;
  configListTitle: string;
  listTitle: string;
  listUrl: string;
  listGuid: string;
  spHttpClient: SPHttpClient;
  aadHttpClientFactory: AadHttpClientFactory;
  userEmail: string;
  onDismiss: () => void;
}

interface IFlowSettingsPanelState {
  loading: boolean;
  saving: boolean;
  canEdit: boolean;
  editing: boolean;
  showingForm: boolean;
  confirmDelete: boolean;
  countEdited: boolean;
  draft?: IConfigDraft;
  legacySite?: IConfigDraft;
  columns: IFormulaColumn[];
  environments: IFlowEnvironment[];
  environmentId: string;
  connections: IFlowConnection[];
  connectionName: string;
  connectionsLoading: boolean;
  flowHint: string;
  consentUrl: string;
  errorMessage: string;
  warningMessage: string;
  statusMessage: string;
}

const RESET_OPTIONS: IDropdownOption[] = [
  { key: 'None', text: 'Never' },
  { key: 'Day', text: 'Every day' },
  { key: 'Month', text: 'Every month' },
  { key: 'Year', text: 'Every year' }
];

export default class FlowSettingsPanel
  extends React.Component<IFlowSettingsPanelProps, IFlowSettingsPanelState> {

  private _connectionRequest = 0;

  public constructor(props: IFlowSettingsPanelProps) {
    super(props);
    this.state = {
      loading: true,
      saving: false,
      canEdit: false,
      editing: false,
      showingForm: false,
      confirmDelete: false,
      countEdited: false,
      columns: [],
      environments: [],
      environmentId: '',
      connections: [],
      connectionName: '',
      connectionsLoading: false,
      flowHint: '',
      consentUrl: '',
      errorMessage: '',
      warningMessage: '',
      statusMessage: ''
    };
  }

  public componentDidMount(): void {
    this._load(true).catch((error: unknown) => {
      this.setState({ loading: false, errorMessage: messageOf(error) });
    });
  }

  public render(): React.ReactElement {
    const draft = this.state.draft;
    const showRow = !!draft && draft.id !== undefined && !this.state.showingForm;
    const showAddLink = !!draft && draft.id === undefined && !this.state.showingForm;
    return (
      <Dialog
        hidden={false}
        onDismiss={this.props.onDismiss}
        minWidth={280}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'AutoGen Flow 1.20',
          showCloseButton: true
        }}
        modalProps={{
          isBlocking: false,
          containerClassName: styles.dialogHost
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
        {this.state.consentUrl && (
          <button type="button" className={styles.addLink} onClick={this._openConsent}>
            {this.state.consentUrl.indexOf('make.powerautomate.com') >= 0 ? 'Open Power Automate' : 'Sign in to the SharePoint connection'}
          </button>
        )}
        {this.state.statusMessage && (
          <div className={styles.status}>
            <MessageBar messageBarType={MessageBarType.success}>{this.state.statusMessage}</MessageBar>
          </div>
        )}
        {!this.state.loading && showAddLink && (
          <button type="button" className={styles.addLink} onClick={this._onAdd} disabled={!this.state.canEdit}>
            Add AutoGen Configuration
          </button>
        )}
        {showRow && draft && this._renderSavedRow(draft)}
        {!this.state.loading && draft && this.state.showingForm && this._renderForm(draft)}
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
          <span>Modified</span>
          <span>Modified by</span>
          <span />
        </div>
        <div className={styles.dataRow}>
          <span className={styles.titleLink}>{draft.title || this.props.listTitle}</span>
          <span>{this._columnTitle(draft.numberColumnInternalName)}</span>
          <span>{draft.formula}</span>
          <span>{this._rowStatus(draft)}</span>
          <span>{formatModified(draft.modified)}</span>
          <span>{draft.modifiedBy || 'Not saved yet'}</span>
          <IconButton
            iconProps={{ iconName: 'More' }}
            ariaLabel="AutoGen actions"
            menuProps={{ items: this._rowItems() }}
            disabled={this.state.saving}
          />
        </div>
        {this.state.confirmDelete && (
          <div className={styles.confirm}>
            <span>Delete AutoGen from this list?</span>
            <PrimaryButton text={this.state.saving ? 'Deleting' : 'Delete'} onClick={this._onDelete} disabled={this.state.saving} />
            <DefaultButton text="Cancel" onClick={this._onCancelDelete} disabled={this.state.saving} />
          </div>
        )}
        {this._openFlowLink()}
      </div>
    );
  }

  private _renderForm(draft: IConfigDraft): React.ReactElement {
    const editing = this.state.editing && this.state.canEdit;
    const busy = this.state.saving || !editing;
    const options = numberColumnChoices(this.state.columns).map((column) => ({
      key: column.internalName,
      text: column.title
    }));
    const inserts = [
      { key: '{counter}', text: 'Counter {counter}' },
      { key: '{counter:4}', text: 'Counter, 4 digits {counter:4}' },
      { key: '{yyyy}', text: 'Year {yyyy}' },
      { key: '{yy}', text: 'Year, 2 digits {yy}' },
      { key: '{MM}', text: 'Month {MM}' },
      { key: '{dd}', text: 'Day {dd}' },
      ...formulaInserts(this.state.columns, draft.numberColumnInternalName).map((column) => ({
        key: column.token,
        text: `${column.title} ${column.token}`
      }))
    ];
    return (
      <div className={styles.panel}>
        <p className={styles.intro}>
          Choose the environment. Saving an active rule creates this account’s SharePoint connection, then creates the numbering flow for this list.
        </p>
        {this._flowPickers()}
        {!this.state.canEdit && (
          <div className={styles.status}>
            <MessageBar messageBarType={MessageBarType.info}>
              You can view this rule. Edit permission on the master list is required to change it.
            </MessageBar>
          </div>
        )}
        <Dropdown
          label="Number column"
          required={true}
          placeholder={options.length === 0 ? 'Add a single line of text column first' : 'Select a column'}
          options={options}
          selectedKey={draft.numberColumnInternalName || undefined}
          disabled={busy || options.length === 0}
          onChange={this._onColumnChange}
        />
        <TextField
          label="Formula"
          required={true}
          description="Use {counter} or {counter:4}. Dates: {yyyy} {MM} {dd}. Columns: {Title}."
          multiline
          rows={2}
          value={draft.formula}
          disabled={busy}
          onChange={this._onFormulaChange}
        />
        <Dropdown
          label="Add to the formula"
          placeholder="Insert a token"
          options={inserts}
          selectedKey={undefined}
          disabled={busy}
          onChange={this._onInsert}
        />
        <TextField
          label="Pad length"
          description="Width for {counter}. {counter:4} uses its own width. Use 0 through 12."
          value={String(draft.padLength)}
          disabled={busy}
          onChange={this._onPadChange}
        />
        <Dropdown
          label="Reset period"
          options={RESET_OPTIONS}
          selectedKey={draft.resetPeriod}
          disabled={busy}
          onChange={this._onResetChange}
        />
        <TextField
          label="Last number used"
          description="The next item adds one. Use 0 to start at 1."
          value={String(draft.currentCount)}
          disabled={busy}
          onChange={this._onCountChange}
        />
        <Checkbox label="Active" checked={draft.isActive} disabled={busy} onChange={this._onActiveChange} />
        <p className={styles.preview}>{this._preview(draft)}</p>
        <div className={styles.actions}>
          {editing && (
            <PrimaryButton
              text={this.state.saving ? 'Saving' : 'Save settings'}
              onClick={this._onSave}
              disabled={this.state.saving || this.state.connectionsLoading}
            />
          )}
          {draft.id !== undefined && (
            <DefaultButton text="Manage flow" onClick={this._onManageFlow} disabled={this.state.saving} />
          )}
          {draft.id !== undefined && (
            <DefaultButton text="Back" onClick={this._onBack} disabled={this.state.saving} />
          )}
          <DefaultButton text="Close" onClick={this.props.onDismiss} disabled={this.state.saving} />
        </div>
      </div>
    );
  }

  private _openFlowLink(): React.ReactElement {
    return (
      <button type="button" className={styles.addLink} onClick={this._onManageFlow} disabled={this.state.saving}>
        Manage flow
      </button>
    );
  }

  private _onManageFlow = (): void => {
    this._manageFlow().catch((error: unknown) => {
      this.setState({ saving: false, errorMessage: messageOf(error), statusMessage: '' });
    });
  };

  private async _manageFlow(): Promise<void> {
    const draft = this.state.draft;
    if (!draft || draft.id === undefined) {
      this.setState({
        warningMessage: 'Save the rule first. Manage flow opens the numbering flow.',
        errorMessage: '',
        statusMessage: ''
      });
      return;
    }
    let flowId = draft.flowId;
    let environmentId = draft.flowEnvironment || this.state.environmentId;
    if (!flowId || draft.flowState !== 'Started') {
      this.setState({ saving: true, errorMessage: '', warningMessage: '', statusMessage: '', consentUrl: '' });
      const synced = await syncListFlow(this.props.aadHttpClientFactory, this._target(), {
        flowId: draft.flowId,
        flowEnvironment: environmentId,
        flowConnection: this.state.connectionName || draft.flowConnection,
        flowState: draft.flowState,
        note: '',
        consentUrl: ''
      }, true);
      await saveFlowBinding(this.props.spHttpClient, this.props.siteAbsoluteUrl, this._listTitle(), draft.id, {
        flowId: synced.flowId,
        flowEnvironment: synced.flowEnvironment,
        flowConnection: synced.flowConnection,
        flowState: synced.flowState
      });
      flowId = synced.flowId;
      environmentId = synced.flowEnvironment;
      this.setState({
        saving: false,
        consentUrl: synced.consentUrl,
        warningMessage: synced.note,
        statusMessage: synced.flowState === 'Started' ? 'Numbering is on.' : 'Saved the rule for this list.',
        draft: {
          ...draft,
          isActive: true,
          flowId: synced.flowId,
          flowEnvironment: synced.flowEnvironment,
          flowConnection: synced.flowConnection,
          flowState: synced.flowState
        }
      });
    }
    if (flowId && environmentId) {
      window.open(flowPortalUrl(environmentId, flowId), '_blank', 'noopener,noreferrer');
    }
  };

  private _openConsent = (): void => {
    if (!this.state.consentUrl) {
      return;
    }
    window.open(this.state.consentUrl, '_blank', 'noopener,noreferrer');
  };

  private _flowPickers(): React.ReactElement {
    const environmentKey = this.state.environments.some((item) => item.id === this.state.environmentId)
      ? this.state.environmentId
      : undefined;
    const locked = !this.state.canEdit || this.state.saving || this.state.connectionsLoading;
    return (
      <div className={styles.pickers}>
        <Dropdown
          label="Environment"
          placeholder={this.state.connectionsLoading ? 'Loading environments' : 'Choose an environment'}
          options={this.state.environments.map((item) => ({ key: item.id, text: item.displayName }))}
          selectedKey={environmentKey}
          disabled={locked || this.state.environments.length === 0}
          onChange={this._onEnvironmentChange}
        />
        {this.state.flowHint && <p className={styles.intro}>{this.state.flowHint}</p>}
      </div>
    );
  }

  private _statusLine(): string {
    const draft = this.state.draft;
    if (this.state.loading) {
      return 'Loading settings.';
    }
    if (!draft || draft.id === undefined) {
      return 'Not saved yet.';
    }
    const who = draft.modifiedBy ? ` by ${draft.modifiedBy}` : '';
    const when = draft.modified ? `Last modified ${formatModified(draft.modified)}${who}.` : 'Saved for this list.';
    const flow = draft.flowState === 'Started' ? ' Numbering is on.' : '';
    return `${when}${flow}`;
  }

  private _rowStatus(draft: IConfigDraft): string {
    if (!draft.isActive) {
      return 'Off';
    }
    return draft.flowState === 'Started' ? 'Active' : 'Saved';
  }

  private _settingsItems(): IContextualMenuItem[] {
    const editing = this.state.editing && this.state.canEdit;
    return [
      {
        key: 'edit',
        text: 'Edit',
        iconProps: { iconName: 'Edit' },
        disabled: !this.state.canEdit || editing || this.state.saving || !this.state.draft || this.state.draft.id === undefined,
        onClick: this._onEdit
      },
      {
        key: 'view',
        text: 'View',
        iconProps: { iconName: 'View' },
        disabled: !this.state.draft || this.state.draft.id === undefined || this.state.saving,
        onClick: this._onView
      },
      {
        key: 'manage',
        text: 'Manage flow',
        iconProps: { iconName: 'OpenInNewWindow' },
        disabled: this.state.saving || !this.state.draft || this.state.draft.id === undefined,
        onClick: this._onManageFlow
      }
    ];
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
        key: 'view',
        text: 'View',
        iconProps: { iconName: 'View' },
        disabled: this.state.saving,
        onClick: this._onView
      },
      {
        key: 'open',
        text: 'Manage flow',
        iconProps: { iconName: 'OpenInNewWindow' },
        disabled: this.state.saving || !this.state.draft || this.state.draft.id === undefined,
        onClick: this._onManageFlow
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

  private _columnTitle(internalName: string): string {
    const match = this.state.columns.filter((column) => column.internalName === internalName)[0];
    return match ? match.title : internalName;
  }

  private _preview(draft: IConfigDraft): string {
    try {
      return planNextCode({
        formula: draft.formula,
        resetPeriod: draft.resetPeriod,
        lastResetDate: draft.lastResetDate,
        currentCount: draft.currentCount,
        padLength: draft.padLength,
        now: new Date()
      }).code;
    } catch (error) {
      return messageOf(error);
    }
  }

  private _onAdd = (): void => {
    this.setState({ showingForm: true, editing: true, confirmDelete: false, statusMessage: '', errorMessage: '', warningMessage: '' });
  };

  private _onEdit = (): void => {
    this.setState({ showingForm: true, editing: true, confirmDelete: false, statusMessage: '', errorMessage: '', warningMessage: '' });
  };

  private _onView = (): void => {
    this.setState({ showingForm: true, editing: false, confirmDelete: false, statusMessage: '', errorMessage: '', warningMessage: '' });
  };

  private _onBack = (): void => {
    this.setState({ showingForm: false, editing: false, confirmDelete: false });
  };

  private _onAskDelete = (): void => {
    this.setState({ confirmDelete: true, errorMessage: '', warningMessage: '' });
  };

  private _onCancelDelete = (): void => {
    this.setState({ confirmDelete: false });
  };

  private _onDelete = (): void => {
    this._delete().catch((error: unknown) => {
      this.setState({ saving: false, errorMessage: messageOf(error) });
    });
  };

  private _onSave = (): void => {
    this._save().catch((error: unknown) => {
      this.setState({ saving: false, errorMessage: messageOf(error) });
    });
  };

  private _onEnvironmentChange = (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (!option) {
      return;
    }
    const environmentId = String(option.key);
    const environment = this.state.environments.filter((item) => item.id === environmentId)[0];
    if (!environment) {
      return;
    }
    const request = this._connectionRequest + 1;
    this._connectionRequest = request;
    this.setState({ environmentId, connectionsLoading: true, connectionName: '', connections: [], flowHint: '' });
    listConnections(this.props.aadHttpClientFactory, environment)
      .then((connections) => {
        if (request !== this._connectionRequest) {
          return;
        }
        const chosen = chooseConnection(connections, this.props.userEmail, this.state.draft ? this.state.draft.flowConnection : '');
        this.setState({
          connectionsLoading: false,
          connections,
          connectionName: chosen ? chosen.name : '',
          flowHint: chosen ? '' : 'Save settings. The app creates this account’s SharePoint connection, then creates the numbering flow.'
        });
      })
      .catch((error: unknown) => {
        if (request !== this._connectionRequest) {
          return;
        }
        this.setState({ connectionsLoading: false, connections: [], connectionName: '', flowHint: messageOf(error) });
      });
  };

  private _onColumnChange = (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (!option) {
      return;
    }
    this._patch({ numberColumnInternalName: String(option.key) });
  };

  private _onFormulaChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this._patch({ formula: value || '' });
  };

  private _onInsert = (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (!option || !this.state.draft) {
      return;
    }
    const token = String(option.key);
    const formula = this.state.draft.formula || '';
    this._patch({ formula: formula.indexOf(token) >= 0 ? formula : `${formula}${token}` });
  };

  private _onPadChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    const padLength = Math.floor(Number(value));
    this._patch({ padLength: Number.isFinite(padLength) ? padLength : 0 });
  };

  private _onResetChange = (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (!option) {
      return;
    }
    this._patch({ resetPeriod: option.key as ResetPeriod });
  };

  private _onCountChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this.setState({ countEdited: true });
    this._patch({ currentCount: wholeCount(value || '0') });
  };

  private _onActiveChange = (_event?: React.FormEvent<HTMLElement | HTMLInputElement>, checked?: boolean): void => {
    this._patch({ isActive: checked === true });
  };

  private _patch(patch: Partial<IConfigDraft>): void {
    if (!this.state.draft) {
      return;
    }
    this.setState({ draft: { ...this.state.draft, ...patch } });
  }

  private async _load(reportRuns = false): Promise<void> {
    const listTitle = this._listTitle();
    await ensureConfigList(this.props.spHttpClient, this.props.siteAbsoluteUrl, listTitle);
    await showConfigList(this.props.spHttpClient, this.props.siteAbsoluteUrl, listTitle);
    const [rows, columns, canEdit] = await Promise.all([
      loadConfigRows(this.props.spHttpClient, this.props.siteAbsoluteUrl, listTitle),
      loadFormulaColumns(this.props.spHttpClient, this.props.webAbsoluteUrl, this.props.listGuid),
      canEditConfigList(this.props.spHttpClient, this.props.siteAbsoluteUrl, listTitle)
    ]);
    const existing = findConfigForList(rows, this.props.listGuid);
    const draft = existing || draftForCurrentList(this.props.listTitle, this.props.listUrl, this.props.listGuid);
    this.setState({
      loading: false,
      canEdit,
      columns,
      draft,
      legacySite: findSiteFlow(rows),
      countEdited: false,
      confirmDelete: false,
      showingForm: false,
      editing: false,
      errorMessage: ''
    });
    if (canEdit) {
      this._loadFlowChoices(draft.flowEnvironment, draft.flowConnection).catch((error: unknown) => {
        this.setState({ connectionsLoading: false, flowHint: messageOf(error) });
      });
    }
    if (reportRuns) {
      this._loadFlowProblem(draft).catch((error: unknown) => {
        this.setState({ warningMessage: messageOf(error) });
      });
    }
  }

  private async _loadFlowProblem(draft: IConfigDraft): Promise<void> {
    if (!draft.flowId || draft.flowState !== 'Started') {
      return;
    }
    const problem = await readFlowProblem(this.props.aadHttpClientFactory, draft.flowEnvironment, draft.flowId);
    if (problem) {
      this.setState({ warningMessage: problem });
    }
  }

  private async _loadFlowChoices(preferredEnvironment: string, preferredConnection: string): Promise<void> {
    this.setState({ connectionsLoading: true, flowHint: '' });
    const environments = await listEnvironments(this.props.aadHttpClientFactory);
    const environment = environments.filter((item) => item.id === preferredEnvironment)[0]
      || environments.filter((item) => item.id === this.state.environmentId)[0]
      || environments[0];
    if (!environment) {
      this.setState({
        connectionsLoading: false,
        environments,
        flowHint: 'Power Automate did not return an environment for this account.'
      });
      return;
    }
    const connections = await listConnections(this.props.aadHttpClientFactory, environment);
    const chosen = chooseConnection(connections, this.props.userEmail, preferredConnection);
    this.setState({
      connectionsLoading: false,
      environments,
      environmentId: environment.id,
      connections,
      connectionName: chosen ? chosen.name : '',
      flowHint: chosen ? '' : 'Save settings. The app creates this account’s SharePoint connection, then creates the numbering flow.'
    });
  }

  private async _save(): Promise<void> {
    const draft = this.state.draft;
    if (!draft || !this.state.canEdit) {
      return;
    }
    const prepared = prepareRule(draft.formula, draft.padLength);
    const next: IConfigDraft = {
      ...draft,
      ...prepared,
      title: draft.title.trim() || this.props.listTitle,
      targetListTitle: this.props.listTitle,
      targetListUrl: this.props.listUrl,
      targetListGuid: this.props.listGuid,
      itemQuery: buildItemQuery(draft.numberColumnInternalName, prepared.formulaColumns)
    };
    const problem = validateDraft(next);
    if (problem) {
      this.setState({ errorMessage: problem, statusMessage: '', warningMessage: '' });
      return;
    }
    this.setState({ saving: true, errorMessage: '', statusMessage: '', warningMessage: '', consentUrl: '' });
    const listTitle = this._listTitle();
    const id = await saveConfigRow(
      this.props.spHttpClient,
      this.props.siteAbsoluteUrl,
      listTitle,
      next,
      next.id === undefined || this.state.countEdited,
      false
    );
    await ensureTargetNumberColumn(
      this.props.spHttpClient,
      this.props.listUrl,
      this.props.listGuid,
      next.numberColumnInternalName
    );
    let saved: IConfigDraft = { ...next, id };
    let warningMessage = '';
    let consentUrl = '';
    let statusMessage = saved.isActive
      ? 'Saved. Numbering is on. Add a new item and leave the number column empty. For an item that is already blank, open it and save it once.'
      : 'Saved the settings for this list. Numbering is turned off.';
    try {
      const synced = await syncListFlow(this.props.aadHttpClientFactory, this._target(), {
        flowId: saved.flowId,
        flowEnvironment: this.state.environmentId || saved.flowEnvironment,
        flowConnection: this.state.connectionName || saved.flowConnection,
        flowState: saved.flowState,
        note: '',
        consentUrl: ''
      }, saved.isActive);
      await saveFlowBinding(this.props.spHttpClient, this.props.siteAbsoluteUrl, listTitle, id, {
        flowId: synced.flowId,
        flowEnvironment: synced.flowEnvironment,
        flowConnection: synced.flowConnection,
        flowState: synced.flowState
      });
      saved = {
        ...saved,
        flowId: synced.flowId,
        flowEnvironment: synced.flowEnvironment,
        flowConnection: synced.flowConnection,
        flowState: synced.flowState
      };
      warningMessage = synced.note;
      consentUrl = synced.consentUrl;
      if (saved.isActive && synced.flowState !== 'Started') {
        statusMessage = 'Saved the rule for this list.';
      }
      if (synced.flowState === 'Started') {
        const legacyNote = await this._stopLegacy();
        warningMessage = warningMessage || legacyNote;
      }
    } catch (error) {
      warningMessage = messageOf(error);
      statusMessage = 'Saved the rule for this list.';
    }
    this.setState({ saving: false, statusMessage, warningMessage, consentUrl, draft: saved });
    await this._load();
  }

  private async _delete(): Promise<void> {
    const draft = this.state.draft;
    if (!draft || draft.id === undefined || !this.state.canEdit) {
      return;
    }
    this.setState({ saving: true, errorMessage: '', statusMessage: '', warningMessage: '' });
    let warningMessage = '';
    try {
      await removeListFlow(this.props.aadHttpClientFactory, draft.flowId, draft.flowEnvironment);
    } catch (error) {
      warningMessage = `${messageOf(error)} The list rule will still be removed.`;
    }
    await deleteConfigRow(this.props.spHttpClient, this.props.siteAbsoluteUrl, this._listTitle(), draft.id);
    this.setState({
      saving: false,
      statusMessage: 'Removed AutoGen from this list.',
      warningMessage
    });
    await this._load();
  }

  private async _stopLegacy(): Promise<string> {
    const legacy = this.state.legacySite;
    if (!legacy || !legacy.flowId || legacy.flowState === 'Stopped') {
      return '';
    }
    try {
      const environments = await listEnvironments(this.props.aadHttpClientFactory);
      const environment = environments.filter((item) => item.id === legacy.flowEnvironment)[0];
      if (!environment) {
        return 'The previous site numbering flow is still on. Turn it off in Power Automate so each item is numbered once.';
      }
      await setRemoteFlowState(this.props.aadHttpClientFactory, environment, legacy.flowId, 'Stopped');
      await saveConfigRow(
        this.props.spHttpClient,
        this.props.siteAbsoluteUrl,
        this._listTitle(),
        { ...legacy, flowState: 'Stopped' },
        false,
        true
      );
      return '';
    } catch {
      return 'The previous site numbering flow is still on. Turn it off in Power Automate so each item is numbered once.';
    }
  }

  private _target(): IListFlowTarget {
    return {
      webUrl: this.props.webAbsoluteUrl,
      siteUrl: this.props.siteAbsoluteUrl,
      listGuid: this.props.listGuid,
      listTitle: this.props.listTitle,
      configListTitle: this._listTitle()
    };
  }

  private _listTitle(): string {
    return this.props.configListTitle || DEFAULT_CONFIG_LIST_TITLE;
  }
}

function messageOf(error: unknown): string {
  if (error instanceof FlowConsentError) {
    return error.message;
  }
  const text = error instanceof Error ? error.message : 'The rule could not be saved.';
  if (text.toLowerCase().indexOf('no http resource') >= 0) {
    return 'Power Automate did not return a SharePoint connection for this account. Add one in the default environment, then save again.';
  }
  return text;
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
