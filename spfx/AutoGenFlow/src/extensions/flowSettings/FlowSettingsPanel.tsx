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
  siteFlowDraft,
  validateDraft,
  wholeCount
} from '../../services/contract';
import {
  canEditConfigList,
  deleteConfigRow,
  ensureConfigList,
  ensureTargetNumberColumn,
  IFlowBinding,
  loadConfigRows,
  loadFormulaColumns,
  saveConfigRow,
  saveFlowBinding,
  showConfigList
} from '../../services/configListClient';
import { chooseConnection, FlowConsentError, flowPortalUrl, IFlowConnection, IFlowEnvironment } from '../../services/flowModel';
import { ISyncedFlow, listConnections, listEnvironments, readFlowProblem, setRemoteFlowState, syncListFlow } from '../../services/flowClient';
import { formulaInserts, IFormulaColumn, numberColumnChoices, planNextCode, prepareRule } from '../../services/formula';
import { buildItemQuery } from '../../services/itemQuery';
import { beginPowerAutomateAccess, openPowerAutomateSignIn, showPowerAutomatePage } from '../../services/flowPortal';
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
  private _makerTab: Window | undefined;

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
          title: 'AutoGen Flow 1.34',
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
          Save settings opens Power Automate in a new tab so this account can sign in. One numbering flow for this site collection reads AutoGenFlowConfiguration and writes the number on each related list.
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
    beginPowerAutomateAccess();
    const draft = this.state.draft;
    if (!draft || draft.id === undefined) {
      this.setState({
        warningMessage: 'Save the rule first. Manage flow opens the site numbering flow.',
        errorMessage: '',
        statusMessage: ''
      });
      return;
    }
    this._makerTab = openPowerAutomateSignIn(this.state.environmentId || draft.flowEnvironment);
    this._manageFlow().catch((error: unknown) => {
      this.setState({ saving: false, errorMessage: messageOf(error), statusMessage: '' });
    });
  };

  private async _manageFlow(): Promise<void> {
    const draft = this.state.draft;
    if (!draft || draft.id === undefined) {
      return;
    }
    this.setState({ saving: true, errorMessage: '', warningMessage: '', statusMessage: '', consentUrl: '' });
    const synced = await this._syncShared(draft, draft.id);
    const flowId = synced.flowId;
    const environmentId = synced.flowEnvironment;
    const opened = this._openMessage(synced.flowState === 'Started', synced.note);
    this.setState({
      saving: false,
      consentUrl: synced.consentUrl,
      warningMessage: opened.warningMessage,
      statusMessage: opened.statusMessage,
      draft: {
        ...draft,
        flowId: synced.flowId,
        flowEnvironment: synced.flowEnvironment,
        flowConnection: synced.flowConnection,
        flowState: synced.flowState
      }
    });
    if (flowId && environmentId) {
      showPowerAutomatePage(this._makerTab, flowPortalUrl(environmentId, flowId));
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
    const shared = this.state.legacySite;
    const started = draft.flowState === 'Started' || (!!shared && shared.flowState === 'Started');
    return started ? 'Active' : 'Saved';
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
    beginPowerAutomateAccess();
    const draft = this.state.draft;
    this._makerTab = openPowerAutomateSignIn(this.state.environmentId || (draft && draft.flowEnvironment) || '');
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
          flowHint: chosen ? '' : 'Save settings. The app creates this account’s SharePoint connection, then creates one numbering flow for this site collection.'
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
      flowHint: chosen ? '' : 'Save settings. The app creates this account’s SharePoint connection, then creates one numbering flow for this site collection.'
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
      ? 'No numbering flow was created. The Power Automate tab lists the flows in this environment.'
      : 'Saved the settings for this list. This list is left out of the site numbering flow.';
    try {
      const synced = await this._syncShared(saved, id);
      saved = {
        ...saved,
        flowId: synced.flowId,
        flowEnvironment: synced.flowEnvironment,
        flowConnection: synced.flowConnection,
        flowState: synced.flowState
      };
      consentUrl = synced.consentUrl;
      const opened = this._openMessage(saved.isActive && synced.flowState === 'Started', synced.note);
      statusMessage = saved.isActive ? opened.statusMessage : 'Saved the settings for this list. This list is left out of the site numbering flow.';
      warningMessage = saved.isActive ? opened.warningMessage : '';
      if (saved.flowId && saved.flowEnvironment) {
        showPowerAutomatePage(this._makerTab, flowPortalUrl(saved.flowEnvironment, saved.flowId));
      }
    } catch (error) {
      const opened = this._openMessage(false, messageOf(error));
      warningMessage = opened.warningMessage;
      statusMessage = opened.statusMessage;
    }
    this.setState({ saving: false, statusMessage, warningMessage, consentUrl, draft: saved });
    await this._load();
  }

  private _openMessage(started: boolean, note: string): { statusMessage: string; warningMessage: string } {
    if (started) {
      return {
        statusMessage: 'Saved. The site numbering flow is on. A blank number is filled within about a minute.',
        warningMessage: ''
      };
    }
    if (!note || note === 'Power Automate is open.' || note === 'No numbering flow was created.' || note.indexOf('Sign in on the Power Automate tab') === 0) {
      return {
        statusMessage: 'No numbering flow was created. The Power Automate tab lists the flows in this environment.',
        warningMessage: ''
      };
    }
    if (note === 'Turn the numbering flow on in the Power Automate tab.') {
      return { statusMessage: note, warningMessage: '' };
    }
    return {
      statusMessage: 'No numbering flow was created. The Power Automate tab lists the flows in this environment.',
      warningMessage: note
    };
  }

  private _sharedSeed(draft: IConfigDraft): ISyncedFlow {
    const legacy = this.state.legacySite;
    const flowId = (legacy && legacy.flowId) || draft.flowId;
    return {
      flowId,
      flowEnvironment: this.state.environmentId || (legacy && legacy.flowEnvironment) || draft.flowEnvironment,
      flowConnection: this.state.connectionName || (legacy && legacy.flowConnection) || draft.flowConnection,
      flowState: flowId && legacy && legacy.flowId === flowId ? legacy.flowState : draft.flowState,
      note: '',
      consentUrl: ''
    };
  }

  private async _syncShared(draft: IConfigDraft, listItemId: number): Promise<ISyncedFlow> {
    const seed = this._sharedSeed(draft);
    const synced = await syncListFlow(
      this.props.aadHttpClientFactory,
      this._target(),
      seed,
      true
    );
    const binding: IFlowBinding = {
      flowId: synced.flowId,
      flowEnvironment: synced.flowEnvironment,
      flowConnection: synced.flowConnection,
      flowState: synced.flowState
    };
    await saveFlowBinding(this.props.spHttpClient, this.props.siteAbsoluteUrl, this._listTitle(), listItemId, binding);
    await this._storeSharedFlow(binding);
    if (synced.flowState === 'Started') {
      const replaced = await this._stopReplacedFlow(draft.flowId, draft.flowEnvironment, synced.flowId);
      if (replaced && !synced.note) {
        synced.note = replaced;
      }
    }
    return synced;
  }

  private async _storeSharedFlow(binding: IFlowBinding): Promise<void> {
    const listTitle = this._listTitle();
    const legacy = this.state.legacySite;
    if (legacy && legacy.id !== undefined) {
      await saveFlowBinding(this.props.spHttpClient, this.props.siteAbsoluteUrl, listTitle, legacy.id, binding);
      this.setState({
        legacySite: {
          ...legacy,
          flowId: binding.flowId,
          flowEnvironment: binding.flowEnvironment,
          flowConnection: binding.flowConnection,
          flowState: binding.flowState
        }
      });
      return;
    }
    const draft = siteFlowDraft(this.props.siteAbsoluteUrl, this.props.siteTitle);
    draft.flowId = binding.flowId;
    draft.flowEnvironment = binding.flowEnvironment;
    draft.flowConnection = binding.flowConnection;
    draft.flowState = binding.flowState;
    const id = await saveConfigRow(this.props.spHttpClient, this.props.siteAbsoluteUrl, listTitle, draft, false, true);
    this.setState({ legacySite: { ...draft, id } });
  }

  private async _stopReplacedFlow(oldId: string, oldEnvironment: string, keptId: string): Promise<string> {
    if (!oldId || oldId === keptId) {
      return '';
    }
    try {
      const environments = await listEnvironments(this.props.aadHttpClientFactory);
      const environment = environments.filter((item) => item.id === oldEnvironment)[0];
      if (!environment) {
        return 'An older flow for this list is still on. Turn it off in Power Automate so each item is numbered once.';
      }
      await setRemoteFlowState(this.props.aadHttpClientFactory, environment, oldId, 'Stopped');
      return '';
    } catch {
      return 'An older flow for this list is still on. Turn it off in Power Automate so each item is numbered once.';
    }
  }

  private async _delete(): Promise<void> {
    const draft = this.state.draft;
    if (!draft || draft.id === undefined || !this.state.canEdit) {
      return;
    }
    this.setState({ saving: true, errorMessage: '', statusMessage: '', warningMessage: '' });
    await deleteConfigRow(this.props.spHttpClient, this.props.siteAbsoluteUrl, this._listTitle(), draft.id);
    this.setState({
      saving: false,
      statusMessage: 'Removed this list from AutoGen Flow. The site numbering flow stays on for the other lists.',
      warningMessage: ''
    });
    await this._load();
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
