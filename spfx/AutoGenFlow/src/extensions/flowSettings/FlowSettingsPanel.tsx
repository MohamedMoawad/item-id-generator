import * as React from 'react';
import {
  Checkbox,
  DefaultButton,
  Dialog,
  DialogType,
  Dropdown,
  DropdownMenuItemType,
  IContextualMenuItem,
  IDropdownOption,
  IDropdownProps,
  IconButton,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  ResponsiveMode,
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
import { createFlowLogDetail, FlowConsentError, flowEditUrl, flowListUrl, IFlowConnection, IFlowEnvironment } from '../../services/flowModel';
import { listEnvironments, readFlowProblem, syncListFlow } from '../../services/flowClient';
import { buildFlowPowerShell, isFlowUnauthorized } from '../../services/flowScript';
import { flowDisplayName, IListFlowTarget } from '../../services/siteFlow';
import { formulaInserts, IFormulaColumn, isLookupLike, numberColumnChoices, planNextCode, prepareRule } from '../../services/formula';
import { buildItemQuery } from '../../services/itemQuery';
import { fillBlankNumbers } from '../../services/listFill';
import { writeFlowLog } from '../../services/flowLog';
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
  flowScript: string;
  errorMessage: string;
  warningMessage: string;
  statusMessage: string;
}

const OPEN_MENU: IDropdownProps['calloutProps'] = {
  calloutMaxHeight: 360,
  layerProps: {
    styles: {
      root: { zIndex: 10000010 }
    }
  }
};

const RESET_OPTIONS: IDropdownOption[] = [
  { key: 'None', text: 'Never' },
  { key: 'Day', text: 'Every day' },
  { key: 'Month', text: 'Every month' },
  { key: 'Year', text: 'Every year' }
];

export default class FlowSettingsPanel
  extends React.Component<IFlowSettingsPanelProps, IFlowSettingsPanelState> {

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
      flowScript: '',
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
          title: 'AutoGen Flow 1.40',
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
        {this._environmentPicker()}
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
        {this._storedEnvironmentId() && this._openFlowLink()}
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
    const columns = columnMenu(this.state.columns, draft.numberColumnInternalName);
    const inserts = [
      { key: '{counter}', text: 'Counter {counter}' },
      { key: '{counter:4}', text: 'Counter, 4 digits {counter:4}' },
      { key: '{yyyy}', text: 'Year {yyyy}' },
      { key: '{yy}', text: 'Year, 2 digits {yy}' },
      { key: '{MM}', text: 'Month {MM}' },
      { key: '{dd}', text: 'Day {dd}' }
    ];
    return (
      <div className={styles.panel}>
        <p className={styles.intro}>
          Set the number column and the formula, then save. This creates an automated cloud flow for this list. The trigger is When an item is created. Each step is stored in AutoGenFlowLog.
        </p>
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
          responsiveMode={ResponsiveMode.large}
          calloutProps={OPEN_MENU}
          onChange={this._onColumnChange}
        />
        <TextField
          label="Formula"
          required={true}
          description="Use {counter} or {counter:4}. Dates: {yyyy} {MM} {dd}. A lookup column is inserted as {Department.title}."
          multiline
          rows={2}
          value={draft.formula}
          disabled={busy}
          onChange={this._onFormulaChange}
        />
        <Dropdown
          label="Columns on this list"
          placeholder={columns.length === 0 ? 'No columns found on this list' : 'Select a column, including lookups'}
          options={columns}
          selectedKey={undefined}
          disabled={busy || columns.length === 0}
          responsiveMode={ResponsiveMode.large}
          calloutProps={OPEN_MENU}
          onChange={this._onInsert}
        />
        <Dropdown
          label="Counter and date"
          placeholder="Insert a counter or date"
          options={inserts}
          selectedKey={undefined}
          disabled={busy}
          responsiveMode={ResponsiveMode.large}
          calloutProps={OPEN_MENU}
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
          responsiveMode={ResponsiveMode.large}
          calloutProps={OPEN_MENU}
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
              disabled={this.state.saving}
            />
          )}
          <DefaultButton text="Edit flow" onClick={this._onEditFlow} disabled={this.state.saving || !this._storedEnvironmentId()} />
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
      <button type="button" className={styles.addLink} onClick={this._onEditFlow} disabled={this.state.saving}>
        Edit flow
      </button>
    );
  }

  private _environmentPicker(): React.ReactElement {
    const options = this.state.environments.map((environment) => ({
      key: environment.id,
      text: environment.isDefault ? `${environment.displayName} (default)` : environment.displayName
    }));
    return (
      <div className={styles.pickers}>
        <Dropdown
          label="Environment"
          placeholder={options.length === 0 ? 'No Power Automate environment was returned' : 'Select an environment'}
          options={options}
          selectedKey={this.state.environmentId || undefined}
          disabled={this.state.loading || this.state.saving || options.length === 0}
          responsiveMode={ResponsiveMode.large}
          calloutProps={OPEN_MENU}
          onChange={this._onEnvironmentChange}
        />
        <div className={styles.actions}>
          <DefaultButton
            text="Edit flow"
            onClick={this._onEditFlow}
            disabled={this.state.saving || !this._storedEnvironmentId()}
          />
          {this.state.flowScript && (
            <PrimaryButton text="Create flow with PowerShell" onClick={this._downloadScript} disabled={this.state.saving} />
          )}
        </div>
        {this.state.flowScript && (
          <MessageBar messageBarType={MessageBarType.warning}>
            Power Automate returned HTTP 401, so this page did not create the flow. Run the PowerShell script. It signs you in, creates the flow in the selected environment, and opens Edit flow in a new tab.
          </MessageBar>
        )}
      </div>
    );
  }

  private _onEnvironmentChange = (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (!option) {
      return;
    }
    const environmentId = String(option.key);
    this.setState({ environmentId, flowScript: '' });
    this._patch({ flowEnvironment: environmentId });
  };

  private _onEditFlow = (): void => {
    const environmentId = this._storedEnvironmentId();
    const flowId = this._storedFlowId();
    if (!environmentId) {
      return;
    }
    const url = flowId ? flowEditUrl(environmentId, flowId) : flowListUrl(environmentId);
    window.open(url, '_blank', 'noopener,noreferrer');
  };

  private _downloadScript = (): void => {
    if (!this.state.flowScript) {
      return;
    }
    const blob = new Blob([this.state.flowScript], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'autogen-flow.ps1';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  private _openConsent = (): void => {
    if (!this.state.consentUrl) {
      return;
    }
    window.open(this.state.consentUrl, '_blank', 'noopener,noreferrer');
  };

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
    const flow = draft.isActive ? ' The number is written on this list.' : '';
    return `${when}${flow}`;
  }

  private _rowStatus(draft: IConfigDraft): string {
    return draft.isActive ? 'Active' : 'Off';
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
      ...(this._storedEnvironmentId() ? [{
        key: 'manage',
        text: 'Edit flow',
        iconProps: { iconName: 'OpenInNewWindow' },
        disabled: this.state.saving || !this._storedEnvironmentId(),
        onClick: this._onEditFlow
      }] : [])
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
      ...(this._storedEnvironmentId() ? [{
        key: 'open',
        text: 'Edit flow',
        iconProps: { iconName: 'OpenInNewWindow' },
        disabled: this.state.saving || !this._storedEnvironmentId(),
        onClick: this._onEditFlow
      }] : []),
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
      const detail = messageOf(error);
      this.setState({ saving: false, errorMessage: detail });
      writeFlowLog(this.props.spHttpClient, this.props.siteAbsoluteUrl, {
        step: 'Save rule',
        status: 'Failed',
        detail,
        listTitle: this.props.listTitle,
        listGuid: this.props.listGuid
      }).catch(() => undefined);
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
    if (!option || option.itemType === DropdownMenuItemType.Header || !this.state.draft) {
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
    const [rows, columns, canEdit, environments] = await Promise.all([
      loadConfigRows(this.props.spHttpClient, this.props.siteAbsoluteUrl, listTitle),
      loadFormulaColumns(this.props.spHttpClient, this.props.webAbsoluteUrl, this.props.listGuid),
      canEditConfigList(this.props.spHttpClient, this.props.siteAbsoluteUrl, listTitle),
      this._loadEnvironments()
    ]);
    const existing = findConfigForList(rows, this.props.listGuid);
    const draft = existing || draftForCurrentList(this.props.listTitle, this.props.listUrl, this.props.listGuid);
    const environmentId = this._selectedEnvironment(environments, draft.flowEnvironment);
    this.setState({
      loading: false,
      canEdit,
      columns,
      environments,
      environmentId,
      draft: { ...draft, flowEnvironment: environmentId || draft.flowEnvironment },
      legacySite: findSiteFlow(rows),
      countEdited: false,
      confirmDelete: false,
      showingForm: false,
      editing: false,
      errorMessage: ''
    });
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
    await writeFlowLog(this.props.spHttpClient, this.props.siteAbsoluteUrl, {
      step: 'Save rule',
      status: 'Completed',
      detail: `Column ${saved.numberColumnInternalName}. Formula ${saved.formula}. Active ${saved.isActive ? 'yes' : 'no'}.`,
      listTitle: this.props.listTitle,
      listGuid: this.props.listGuid
    });
    let statusMessage = saved.isActive
      ? 'Saved. The automated flow for this list is being created. Open AutoGenFlowLog to see each step.'
      : 'Saved. This list is off. Open AutoGenFlowLog to see each step.';
    let warningMessage = '';
    let flowScript = '';
    if (saved.isActive) {
      const created = await this._createListFlow(saved);
      saved = created.draft;
      statusMessage = created.statusMessage;
      warningMessage = created.warningMessage;
      flowScript = created.flowScript;
    }
    this.setState({
      saving: false,
      statusMessage,
      warningMessage,
      consentUrl: '',
      flowScript,
      environmentId: saved.flowEnvironment || this.state.environmentId,
      draft: saved,
      showingForm: false
    });
    if (saved.isActive) {
      await fillBlankNumbers({
        spHttpClient: this.props.spHttpClient,
        siteAbsoluteUrl: this.props.siteAbsoluteUrl,
        webAbsoluteUrl: this.props.webAbsoluteUrl,
        configListTitle: listTitle,
        listGuid: this.props.listGuid,
        listTitle: this.props.listTitle
      });
    }
    await this._load();
  }

  private async _createListFlow(draft: IConfigDraft): Promise<{ draft: IConfigDraft; statusMessage: string; warningMessage: string; flowScript: string }> {
    const target: IListFlowTarget = {
      webUrl: this.props.webAbsoluteUrl,
      siteUrl: this.props.siteAbsoluteUrl,
      listGuid: this.props.listGuid,
      listTitle: this.props.listTitle,
      configListTitle: this._listTitle()
    };
    const name = flowDisplayName(this.props.listTitle);
    const environmentId = this.state.environmentId || draft.flowEnvironment;
    try {
      const synced = await syncListFlow(this.props.aadHttpClientFactory, target, {
        flowId: draft.flowId,
        flowEnvironment: environmentId,
        flowConnection: draft.flowConnection,
        flowState: draft.flowState,
        note: '',
        consentUrl: ''
      }, true);
      if (draft.id !== undefined && (synced.flowId || synced.flowEnvironment)) {
        await saveFlowBinding(this.props.spHttpClient, this.props.siteAbsoluteUrl, this._listTitle(), draft.id, {
          flowId: synced.flowId,
          flowEnvironment: synced.flowEnvironment,
          flowConnection: synced.flowConnection,
          flowState: synced.flowState
        });
      }
      const started = synced.flowState === 'Started' && synced.flowId.trim().length > 0;
      const flowScript = isFlowUnauthorized(synced.note) ? buildFlowPowerShell({
        environmentId: synced.flowEnvironment,
        environmentName: this._environmentName(synced.flowEnvironment),
        apiHost: this._environmentHost(synced.flowEnvironment),
        target,
        connectionName: synced.flowConnection
      }) : '';
      const note = flowScript ? `${synced.note} Create flow with PowerShell is ready in AutoGen Flow.` : synced.note;
      const detail = createFlowLogDetail({
        name,
        environmentId: synced.flowEnvironment,
        flowId: synced.flowId,
        flowState: synced.flowState,
        note
      });
      await writeFlowLog(this.props.spHttpClient, this.props.siteAbsoluteUrl, {
        step: 'Create flow',
        status: started ? 'Completed' : 'Failed',
        detail,
        listTitle: this.props.listTitle,
        listGuid: this.props.listGuid
      });
      return {
        draft: {
          ...draft,
          flowId: synced.flowId,
          flowEnvironment: synced.flowEnvironment,
          flowConnection: synced.flowConnection,
          flowState: synced.flowState
        },
        statusMessage: createStatus(started, flowScript.length > 0),
        warningMessage: started ? '' : (synced.note || 'The flow was not created.'),
        flowScript
      };
    } catch (error) {
      const detail = messageOf(error);
      const flowScript = isFlowUnauthorized(detail) ? buildFlowPowerShell({
        environmentId,
        environmentName: this._environmentName(environmentId),
        apiHost: this._environmentHost(environmentId),
        target,
        connectionName: draft.flowConnection
      }) : '';
      await writeFlowLog(this.props.spHttpClient, this.props.siteAbsoluteUrl, {
        step: 'Create flow',
        status: 'Failed',
        detail: createFlowLogDetail({
          name,
          environmentId,
          flowId: draft.flowId,
          flowState: draft.flowState,
          note: flowScript ? `${detail} Create flow with PowerShell is ready in AutoGen Flow.` : detail
        }),
        listTitle: this.props.listTitle,
        listGuid: this.props.listGuid
      });
      return {
        draft,
        statusMessage: flowScript
          ? 'Saved. Power Automate returned HTTP 401. Use Create flow with PowerShell. It opens Edit flow in a new tab.'
          : 'Saved. The flow was not created. Open AutoGenFlowLog and read the Create flow row.',
        warningMessage: detail,
        flowScript
      };
    }
  }

  private async _loadEnvironments(): Promise<IFlowEnvironment[]> {
    try {
      return await listEnvironments(this.props.aadHttpClientFactory);
    } catch (error) {
      this.setState({ warningMessage: messageOf(error) });
      return [];
    }
  }

  private _selectedEnvironment(environments: IFlowEnvironment[], savedId: string): string {
    const saved = savedId.trim();
    if (saved && environments.some((item) => item.id === saved)) {
      return saved;
    }
    return environments.length > 0 ? environments[0].id : saved;
  }

  private _environmentName(environmentId: string): string {
    const match = this.state.environments.filter((item) => item.id === environmentId)[0];
    return match ? match.displayName : 'Default';
  }

  private _environmentHost(environmentId: string): string {
    const match = this.state.environments.filter((item) => item.id === environmentId)[0];
    return match ? match.endpoint : 'https://api.flow.microsoft.com';
  }

  private _storedFlowId(): string {
    const draft = this.state.draft;
    const shared = this.state.legacySite;
    return ((draft && draft.flowId) || (shared && shared.flowId) || '').trim();
  }

  private _storedEnvironmentId(): string {
    const draft = this.state.draft;
    const shared = this.state.legacySite;
    return (this.state.environmentId || (draft && draft.flowEnvironment) || (shared && shared.flowEnvironment) || '').trim();
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
      statusMessage: 'Removed this list from AutoGen Flow.',
      warningMessage: ''
    });
    await this._load();
  }

  private _listTitle(): string {
    return this.props.configListTitle || DEFAULT_CONFIG_LIST_TITLE;
  }
}

function createStatus(started: boolean, scriptReady: boolean): string {
  if (started) {
    return 'Saved. The automated flow is on for this list. Edit flow opens it in a new tab.';
  }
  if (scriptReady) {
    return 'Saved. Power Automate returned HTTP 401. Use Create flow with PowerShell. It opens Edit flow in a new tab.';
  }
  return 'Saved. The flow was not created. Open AutoGenFlowLog and read the Create flow row.';
}

function columnMenu(columns: IFormulaColumn[], numberField: string): IDropdownOption[] {
  const rows = formulaInserts(columns, numberField);
  const lookups = rows.filter((column) => isLookupLike(column.type));
  const plain = rows.filter((column) => !isLookupLike(column.type));
  const options: IDropdownOption[] = [];
  if (lookups.length > 0) {
    options.push({ key: 'lookup-header', text: 'Lookup columns', itemType: DropdownMenuItemType.Header });
    lookups.forEach((column) => options.push({ key: column.token, text: `${column.title} ${column.token}` }));
  }
  if (plain.length > 0) {
    options.push({ key: 'column-header', text: 'Other columns', itemType: DropdownMenuItemType.Header });
    plain.forEach((column) => options.push({ key: column.token, text: `${column.title} ${column.token}` }));
  }
  return options;
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
