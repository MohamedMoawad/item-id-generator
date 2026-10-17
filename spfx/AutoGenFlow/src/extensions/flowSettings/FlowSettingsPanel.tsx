import * as React from 'react';
import {
  DefaultButton,
  Dialog,
  DialogType,
  Dropdown,
  IDropdownOption,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  Spinner,
  SpinnerSize,
  TextField,
  Toggle
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
  ensureConfigList,
  ensureTargetNumberColumn,
  loadConfigRows,
  loadFormulaColumns,
  saveConfigRow,
  showConfigList
} from '../../services/configListClient';
import {
  createRemoteFlow,
  loadFlowContext,
  readRemoteFlow,
  redeployRemoteFlow,
  setRemoteFlowState
} from '../../services/flowClient';
import { FlowConsentError, FlowRunState, FlowStartError, IFlowConnection, IFlowEnvironment } from '../../services/flowModel';
import { formulaInserts, IFormulaColumn, numberColumnChoices, planNextCode, prepareRule } from '../../services/formula';
import { buildItemQuery } from '../../services/itemQuery';
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
  onDismiss: () => void;
}

interface IFlowSettingsPanelState {
  loading: boolean;
  saving: boolean;
  working: boolean;
  canEdit: boolean;
  countEdited: boolean;
  consentRequired: boolean;
  draft?: IConfigDraft;
  columns: IFormulaColumn[];
  errorMessage: string;
  statusMessage: string;
  flowMessage: string;
  numbering: FlowRunState;
  flowId: string;
  siteRowId?: number;
  environment?: IFlowEnvironment;
  connections: IFlowConnection[];
  connectionName: string;
}

const DATE_TOKENS = ['{counter}', '{counter:4}', '{yyyy}', '{yy}', '{MM}', '{dd}'];

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
      working: false,
      canEdit: false,
      countEdited: false,
      consentRequired: false,
      columns: [],
      errorMessage: '',
      statusMessage: '',
      flowMessage: '',
      numbering: 'Missing',
      flowId: '',
      connections: [],
      connectionName: ''
    };
  }

  public componentDidMount(): void {
    this._load().catch((error: unknown) => {
      this.setState({ loading: false, errorMessage: messageOf(error) });
    });
  }

  public render(): React.ReactElement {
    const draft = this.state.draft;
    return (
      <Dialog
        hidden={false}
        onDismiss={this.props.onDismiss}
        minWidth={280}
        maxWidth={520}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'AutoGen',
          subText: `Sequential numbers for ${this.props.listTitle}`,
          showCloseButton: true
        }}
        modalProps={{
          isBlocking: false,
          containerClassName: styles.dialogHost
        }}
      >
        <div className={styles.panel}>
          {this.state.loading && <Spinner size={SpinnerSize.small} label="Loading numbering" />}
          {this.state.errorMessage && <MessageBar messageBarType={MessageBarType.error}>{this.state.errorMessage}</MessageBar>}
          {this.state.statusMessage && <MessageBar messageBarType={MessageBarType.success}>{this.state.statusMessage}</MessageBar>}
          {!this.state.loading && this._flowCard()}
          {draft && !this.state.loading && this._form(draft)}
        </div>
      </Dialog>
    );
  }

  private _flowCard(): React.ReactElement {
    const on = this.state.numbering === 'Started';
    const pillClass = on ? styles.pillOn : (this.state.numbering === 'Stopped' ? styles.pillOff : styles.pillWait);
    const pillText = on ? 'On' : (this.state.numbering === 'Stopped' ? 'Off' : 'Not created');
    return (
      <section className={styles.card}>
        <div className={styles.cardHead}>
          <div>
            <p className={styles.cardTitle}>Site numbering</p>
            <p className={styles.cardCopy}>
              {on
                ? 'Every list on this site collection is numbered from its saved rule.'
                : 'Turn this on once. It covers every list on this site collection.'}
            </p>
          </div>
          <span className={`${styles.pill} ${pillClass}`}>{pillText}</span>
        </div>
        {this.state.consentRequired && (
          <MessageBar messageBarType={MessageBarType.warning}>
            A SharePoint admin approves AutoGen Flow once under Admin center, Advanced, API access.
          </MessageBar>
        )}
        {this.state.flowMessage && <p className={styles.cardCopy}>{this.state.flowMessage}</p>}
        {this.state.connections.length > 1 && (
          <Dropdown
            label="SharePoint connection"
            options={this.state.connections.map((item) => ({ key: item.name, text: item.displayName }))}
            selectedKey={this.state.connectionName || undefined}
            disabled={this.state.working || !this.state.canEdit}
            onChange={this._onConnectionChange}
          />
        )}
        {this.state.connections.length === 1 && (
          <p className={styles.hint}>Connection: {this.state.connections[0].displayName}</p>
        )}
        <div className={styles.actions}>
          {!on && (
            <PrimaryButton
              text={this.state.working ? 'Turning numbering on' : 'Turn on numbering'}
              disabled={this.state.working || this.state.saving || !this.state.canEdit}
              onClick={this._turnOn}
            />
          )}
          {on && (
            <DefaultButton text="Turn off" disabled={this.state.working || !this.state.canEdit} onClick={this._turnOff} />
          )}
          {this.state.flowId && (
            <DefaultButton text="Redeploy" disabled={this.state.working || this.state.saving || !this.state.canEdit} onClick={this._redeploy} />
          )}
        </div>
      </section>
    );
  }

  private _form(draft: IConfigDraft): React.ReactElement {
    const options = numberColumnChoices(this.state.columns).map((column) => ({
      key: column.internalName,
      text: column.title === column.internalName ? column.title : `${column.title} (${column.internalName})`
    }));
    if (draft.numberColumnInternalName && !options.some((option) => option.key === draft.numberColumnInternalName)) {
      options.unshift({ key: draft.numberColumnInternalName, text: draft.numberColumnInternalName });
    }
    const inserts = formulaInserts(this.state.columns, draft.numberColumnInternalName);
    return (
      <section>
        <Dropdown
          label="Number column"
          placeholder={options.length === 0 ? 'Add a single line of text column first' : 'Choose a column'}
          options={options}
          selectedKey={draft.numberColumnInternalName || undefined}
          disabled={!this.state.canEdit || this.state.saving}
          onChange={this._onColumnChange}
        />
        <TextField
          label="Formula"
          description="Example: REQ-{yyyy}{MM}-{counter:4}"
          multiline
          rows={2}
          value={draft.formula}
          disabled={!this.state.canEdit || this.state.saving}
          onChange={this._onFormulaChange}
        />
        <p className={styles.insertLabel}>Insert</p>
        <div className={styles.tokens}>
          {DATE_TOKENS.map((token) => (
            <button key={token} type="button" className={styles.token} data-token={token} onClick={this._onInsertToken}>{token}</button>
          ))}
          {inserts.map((column) => (
            <button key={column.internalName} type="button" className={styles.token} data-token={column.token} onClick={this._onInsertToken}>
              {column.token}
            </button>
          ))}
        </div>
        <TextField
          label="Digits"
          description="Used when the formula says {counter} without a width."
          value={String(draft.padLength)}
          disabled={!this.state.canEdit || this.state.saving}
          onChange={this._onPadChange}
        />
        <Dropdown
          label="Start again"
          options={RESET_OPTIONS}
          selectedKey={draft.resetPeriod}
          disabled={!this.state.canEdit || this.state.saving}
          onChange={this._onResetChange}
        />
        <TextField
          label="Last number used"
          description="The next item adds one. Use 0 to start at 1."
          value={String(draft.currentCount)}
          disabled={!this.state.canEdit || this.state.saving}
          onChange={this._onCountChange}
        />
        <Toggle
          label="This list"
          checked={draft.isActive}
          onText="Numbering"
          offText="Paused"
          disabled={!this.state.canEdit || this.state.saving}
          onChange={this._onActiveChange}
        />
        <div className={styles.preview}>
          <p className={styles.previewLabel}>Next number</p>
          <p className={styles.previewValue}>{this._preview(draft)}</p>
          <p className={styles.hint}>Up to 100 blank items a minute, in the order they were created.</p>
        </div>
        <div className={styles.actions}>
          <PrimaryButton text={this.state.saving ? 'Saving' : 'Save'} disabled={!this.state.canEdit || this.state.saving || this.state.working} onClick={this._onSave} />
          <DefaultButton text="Close" onClick={this.props.onDismiss} />
        </div>
      </section>
    );
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

  private _onInsertToken = (event: React.MouseEvent<HTMLButtonElement>): void => {
    const draft = this.state.draft;
    const token = event.currentTarget.getAttribute('data-token') || '';
    if (!draft || !this.state.canEdit || !token) {
      return;
    }
    this.setState({ draft: { ...draft, formula: `${draft.formula}${token}` }, statusMessage: '' });
  };

  private _onColumnChange = (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    const draft = this.state.draft;
    if (!draft || !option) {
      return;
    }
    this.setState({ draft: { ...draft, numberColumnInternalName: String(option.key) }, statusMessage: '' });
  };

  private _onFormulaChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    const draft = this.state.draft;
    if (!draft) {
      return;
    }
    this.setState({ draft: { ...draft, formula: value || '' }, statusMessage: '' });
  };

  private _onPadChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    const draft = this.state.draft;
    if (!draft) {
      return;
    }
    const padLength = Math.floor(Number(value));
    this.setState({ draft: { ...draft, padLength: Number.isFinite(padLength) ? padLength : 0 }, statusMessage: '' });
  };

  private _onResetChange = (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    const draft = this.state.draft;
    if (!draft || !option) {
      return;
    }
    this.setState({ draft: { ...draft, resetPeriod: option.key as ResetPeriod }, statusMessage: '' });
  };

  private _onCountChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    const draft = this.state.draft;
    if (!draft) {
      return;
    }
    this.setState({
      countEdited: true,
      draft: { ...draft, currentCount: wholeCount(value || '0') },
      statusMessage: ''
    });
  };

  private _onActiveChange = (_event: React.MouseEvent<HTMLElement>, checked?: boolean): void => {
    const draft = this.state.draft;
    if (!draft) {
      return;
    }
    this.setState({ draft: { ...draft, isActive: checked === true }, statusMessage: '' });
  };

  private _onConnectionChange = (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (!option) {
      return;
    }
    this.setState({ connectionName: String(option.key) });
  };

  private _onSave = (): void => {
    this._save(true).catch((error: unknown) => {
      this.setState({ saving: false, errorMessage: messageOf(error) });
    });
  };

  private _turnOn = (): void => {
    this._changeFlow('on').catch((error: unknown) => {
      this._failFlow(error);
    });
  };

  private _turnOff = (): void => {
    this._changeFlow('off').catch((error: unknown) => {
      this._failFlow(error);
    });
  };

  private _redeploy = (): void => {
    this._changeFlow('redeploy').catch((error: unknown) => {
      this._failFlow(error);
    });
  };

  private async _load(): Promise<void> {
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
    const site = findSiteFlow(rows);
    this.setState({
      loading: false,
      canEdit,
      columns,
      draft,
      siteRowId: site && site.id,
      flowId: site ? site.flowId : '',
      connectionName: site ? site.flowConnection : '',
      numbering: site && site.flowState === 'Started' ? 'Started' : (site && site.flowState === 'Stopped' ? 'Stopped' : 'Missing'),
      errorMessage: canEdit ? '' : 'You can view this rule. Edit permission on the master list is required to save it.'
    });
    if (!canEdit) {
      return;
    }
    await this._refreshFlow(
      site ? site.flowId : '',
      site ? site.flowEnvironment : '',
      site ? site.flowConnection : ''
    );
  }

  private async _refreshFlow(flowId: string, environmentId: string, savedConnection: string): Promise<void> {
    try {
      const context = await loadFlowContext(this.props.aadHttpClientFactory);
      const environment = context.environment.id === environmentId || !environmentId
        ? context.environment
        : { ...context.environment, id: environmentId };
      let numbering: FlowRunState = flowId ? 'Stopped' : 'Missing';
      let id = flowId;
      if (flowId) {
        const remote = await readRemoteFlow(this.props.aadHttpClientFactory, environment, flowId);
        numbering = remote.state;
        id = remote.state === 'Missing' ? '' : remote.id;
      }
      const connectionName = savedConnection || (context.connections[0] ? context.connections[0].name : '');
      this.setState({
        consentRequired: false,
        environment,
        connections: context.connections,
        connectionName,
        flowId: id,
        numbering,
        flowMessage: context.connections.length === 0
          ? 'Add a SharePoint connection in Power Automate, then turn numbering on.'
          : ''
      });
    } catch (error) {
      if (error instanceof FlowConsentError) {
        this.setState({ consentRequired: true, flowMessage: '' });
        return;
      }
      this.setState({ flowMessage: messageOf(error) });
    }
  }

  private async _save(announce: boolean): Promise<boolean> {
    const draft = this.state.draft;
    if (!draft) {
      return false;
    }
    const prepared = prepareRule(draft.formula, draft.padLength);
    const next: IConfigDraft = {
      ...draft,
      ...prepared,
      title: draft.title.trim() || this.props.listTitle,
      targetListTitle: this.props.listTitle,
      targetListUrl: this.props.listUrl,
      targetListGuid: this.props.listGuid,
      itemQuery: buildItemQuery(this.props.listGuid, draft.numberColumnInternalName, prepared.formulaColumns)
    };
    const problem = validateDraft(next);
    if (problem) {
      this.setState({ errorMessage: problem, statusMessage: '' });
      return false;
    }
    this.setState({ saving: true, errorMessage: '' });
    const listTitle = this._listTitle();
    await ensureConfigList(this.props.spHttpClient, this.props.siteAbsoluteUrl, listTitle);
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
    const running = this.state.numbering === 'Started';
    this.setState({
      saving: false,
      countEdited: false,
      draft: { ...next, id },
      statusMessage: announce
        ? (running
          ? 'Saved. The next item on this list is numbered within a minute.'
          : 'Saved. Turn on site numbering to start writing numbers.')
        : this.state.statusMessage
    });
    return true;
  }

  private async _changeFlow(action: 'on' | 'off' | 'redeploy'): Promise<void> {
    if (action !== 'off') {
      const saved = await this._save(false);
      if (!saved) {
        return;
      }
    }
    this.setState({ working: true, flowMessage: '', errorMessage: '' });
    const context = this.state.environment
      ? { environment: this.state.environment, connections: this.state.connections }
      : await loadFlowContext(this.props.aadHttpClientFactory);
    const connection = this.state.connectionName || (context.connections[0] ? context.connections[0].name : '');
    if (action !== 'off' && !connection) {
      this.setState({
        working: false,
        environment: context.environment,
        connections: context.connections,
        flowMessage: 'Add a SharePoint connection in Power Automate, then turn numbering on.'
      });
      return;
    }
    const listTitle = this._listTitle();
    let flowId = this.state.flowId;
    let numbering: FlowRunState = this.state.numbering;
    if (action === 'off') {
      if (flowId) {
        await setRemoteFlowState(this.props.aadHttpClientFactory, context.environment, flowId, 'Stopped');
      }
      numbering = 'Stopped';
    } else if (action === 'redeploy' && flowId) {
      const remote = await redeployRemoteFlow(
        this.props.aadHttpClientFactory,
        context.environment,
        flowId,
        this.props.siteAbsoluteUrl,
        listTitle,
        connection
      );
      flowId = remote.id;
      numbering = 'Started';
    } else if (flowId && numbering === 'Stopped') {
      await setRemoteFlowState(this.props.aadHttpClientFactory, context.environment, flowId, 'Started');
      numbering = 'Started';
    } else if (!flowId) {
      const remote = await createRemoteFlow(
        this.props.aadHttpClientFactory,
        context.environment,
        this.props.siteAbsoluteUrl,
        listTitle,
        connection
      );
      flowId = remote.id;
      numbering = 'Started';
    }
    const siteId = await this._persistSite(flowId, context.environment.id, connection, numbering === 'Started' ? 'Started' : 'Stopped');
    this.setState({
      working: false,
      consentRequired: false,
      environment: context.environment,
      connections: context.connections,
      connectionName: connection,
      flowId,
      siteRowId: siteId,
      numbering,
      statusMessage: numbering === 'Started'
        ? 'Numbering is on for every list on this site collection.'
        : 'Numbering is off. Saved rules stay in place.'
    });
  }

  private async _persistSite(flowId: string, environmentId: string, connectionName: string, state: string): Promise<number> {
    const listTitle = this._listTitle();
    const draft: IConfigDraft = {
      ...siteFlowDraft(this.props.siteAbsoluteUrl, this.props.siteTitle),
      id: this.state.siteRowId,
      flowId,
      flowEnvironment: environmentId,
      flowConnection: connectionName,
      flowState: state
    };
    return saveConfigRow(this.props.spHttpClient, this.props.siteAbsoluteUrl, listTitle, draft, false, true);
  }

  private _failFlow(error: unknown): void {
    if (error instanceof FlowConsentError) {
      this.setState({ working: false, saving: false, consentRequired: true, flowMessage: '' });
      return;
    }
    if (error instanceof FlowStartError) {
      this._persistSite(error.flowId, this.state.environment ? this.state.environment.id : '', this.state.connectionName, 'Stopped')
        .then((siteId) => {
          this.setState({
            working: false,
            saving: false,
            flowId: error.flowId,
            siteRowId: siteId,
            numbering: 'Stopped',
            flowMessage: error.message
          });
        })
        .catch(() => {
          this.setState({ working: false, saving: false, flowId: error.flowId, numbering: 'Stopped', flowMessage: error.message });
        });
      return;
    }
    this.setState({ working: false, saving: false, flowMessage: messageOf(error) });
  }

  private _listTitle(): string {
    return this.props.configListTitle || DEFAULT_CONFIG_LIST_TITLE;
  }
}

function messageOf(error: unknown): string {
  if (error instanceof FlowConsentError) {
    return error.message;
  }
  if (error instanceof Error) {
    return error.message;
  }
  return 'The rule could not be saved.';
}
