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
  saveFlowBinding,
  showConfigList
} from '../../services/configListClient';
import {
  createRemoteFlow,
  listConnections,
  loadFlowContext,
  readRemoteFlow,
  redeployRemoteFlow,
  setRemoteFlowState
} from '../../services/flowClient';
import { FlowConsentError, FlowRunState, FlowStartError, IFlowConnection, IFlowEnvironment } from '../../services/flowModel';
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
  onDismiss: () => void;
}

interface IFlowSettingsPanelState {
  loading: boolean;
  saving: boolean;
  working: boolean;
  connectionsLoading: boolean;
  canEdit: boolean;
  countEdited: boolean;
  consentRequired: boolean;
  draft?: IConfigDraft;
  legacySite?: IConfigDraft;
  columns: IFormulaColumn[];
  errorMessage: string;
  statusMessage: string;
  flowMessage: string;
  numbering: FlowRunState;
  flowId: string;
  environments: IFlowEnvironment[];
  environmentId: string;
  connections: IFlowConnection[];
  connectionName: string;
}

interface IPendingFlow {
  environment: IFlowEnvironment;
  connectionName: string;
  itemId: number;
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

  private _rule?: IConfigDraft;
  private _pending?: IPendingFlow;
  private _connectionRequest = 0;

  public constructor(props: IFlowSettingsPanelProps) {
    super(props);
    this.state = {
      loading: true,
      saving: false,
      working: false,
      connectionsLoading: false,
      canEdit: false,
      countEdited: false,
      consentRequired: false,
      columns: [],
      errorMessage: '',
      statusMessage: '',
      flowMessage: '',
      numbering: 'Missing',
      flowId: '',
      environments: [],
      environmentId: '',
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
    const environmentKey = this.state.environments.some((item) => item.id === this.state.environmentId)
      ? this.state.environmentId
      : undefined;
    const connectionKey = this.state.connections.some((item) => item.name === this.state.connectionName)
      ? this.state.connectionName
      : undefined;
    const locked = this.state.working || this.state.connectionsLoading || !this.state.canEdit;
    return (
      <section className={styles.card}>
        <div className={styles.cardHead}>
          <div>
            <p className={styles.cardTitle}>Numbering for this list</p>
            <p className={styles.cardCopy}>
              {on
                ? 'New items on this list receive the next number. The flow reads the rule on the site collection.'
                : 'This list gets its own flow. The flow reads the rule on the site collection.'}
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
        <div className={styles.pickers}>
          <Dropdown
            label="Environment"
            placeholder={this.state.environments.length === 0 ? 'No environment available' : 'Choose an environment'}
            options={this.state.environments.map((item) => ({ key: item.id, text: item.displayName }))}
            selectedKey={environmentKey}
            disabled={locked || this.state.environments.length === 0}
            onChange={this._onEnvironmentChange}
          />
          <Dropdown
            label="SharePoint connection"
            placeholder={this.state.connections.length === 0 ? 'No SharePoint connection' : 'Choose a connection'}
            options={this.state.connections.map((item) => ({ key: item.name, text: item.displayName }))}
            selectedKey={connectionKey}
            disabled={locked || this.state.connections.length === 0}
            onChange={this._onConnectionChange}
          />
        </div>
        <p className={styles.hint}>The connection needs Contribute on this list and on the site collection.</p>
        <div className={styles.actions}>
          {!on && (
            <PrimaryButton
              text={this.state.working ? 'Turning numbering on' : 'Turn on numbering'}
              disabled={locked || this.state.saving}
              onClick={this._turnOn}
            />
          )}
          {on && (
            <DefaultButton text="Turn off" disabled={locked} onClick={this._turnOff} />
          )}
          {this.state.flowId && (
            <DefaultButton text="Redeploy" disabled={locked || this.state.saving} onClick={this._redeploy} />
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
          <p className={styles.hint}>The number is written when the item is created.</p>
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

  private _onEnvironmentChange = (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (!option) {
      return;
    }
    const environmentId = String(option.key);
    if (environmentId === this.state.environmentId) {
      return;
    }
    const environment = this.state.environments.filter((item) => item.id === environmentId)[0];
    if (!environment) {
      return;
    }
    const request = this._connectionRequest + 1;
    this._connectionRequest = request;
    this.setState({ environmentId, connectionsLoading: true, flowMessage: '' });
    listConnections(this.props.aadHttpClientFactory, environment)
      .then((connections) => {
        if (request !== this._connectionRequest) {
          return;
        }
        const current = this.state.connectionName;
        const connectionName = connections.some((item) => item.name === current)
          ? current
          : (connections[0] ? connections[0].name : '');
        this.setState({
          connectionsLoading: false,
          connections,
          connectionName,
          consentRequired: false,
          flowMessage: connections.length === 0
            ? 'Add a SharePoint connection in this environment, then turn numbering on.'
            : ''
        });
      })
      .catch((error: unknown) => {
        if (request !== this._connectionRequest) {
          return;
        }
        if (error instanceof FlowConsentError) {
          this.setState({ connectionsLoading: false, consentRequired: true, connections: [], connectionName: '' });
          return;
        }
        this.setState({ connectionsLoading: false, flowMessage: messageOf(error) });
      });
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
    this._rule = draft;
    const flow = canEdit
      ? await this._readFlow(draft.flowId, draft.flowEnvironment, draft.flowConnection)
      : {};
    const savedNumbering: FlowRunState = draft.flowState === 'Started' ? 'Started' : (draft.flowId ? 'Stopped' : 'Missing');
    this.setState({
      loading: false,
      canEdit,
      columns,
      draft,
      legacySite: findSiteFlow(rows),
      flowId: flow.flowId !== undefined ? flow.flowId : draft.flowId,
      connectionName: flow.connectionName !== undefined ? flow.connectionName : draft.flowConnection,
      environmentId: flow.environmentId !== undefined ? flow.environmentId : draft.flowEnvironment,
      numbering: flow.numbering !== undefined ? flow.numbering : savedNumbering,
      environments: flow.environments || [],
      connections: flow.connections || [],
      consentRequired: flow.consentRequired === true,
      flowMessage: flow.flowMessage || '',
      errorMessage: canEdit ? '' : 'You can view this rule. Edit permission on the master list is required to save it.'
    });
  }

  private async _readFlow(flowId: string, environmentId: string, savedConnection: string): Promise<Partial<IFlowSettingsPanelState>> {
    try {
      const context = await loadFlowContext(this.props.aadHttpClientFactory);
      const matched = context.environments.filter((item) => item.id === environmentId)[0];
      const environment = matched || context.environment;
      const connections = !matched || matched.id === context.environment.id
        ? context.connections
        : await listConnections(this.props.aadHttpClientFactory, matched);
      let numbering: FlowRunState = flowId ? 'Stopped' : 'Missing';
      let id = flowId;
      let flowMessage = connections.length === 0
        ? 'Add a SharePoint connection in this environment, then turn numbering on.'
        : '';
      if (flowId && matched) {
        const remote = await readRemoteFlow(this.props.aadHttpClientFactory, matched, flowId);
        numbering = remote.state;
        id = remote.state === 'Missing' ? '' : remote.id;
      } else if (flowId && !matched) {
        numbering = 'Stopped';
        flowMessage = 'The environment saved on this list is not available to this account. Choose an environment and redeploy.';
      }
      const connectionName = connections.some((item) => item.name === savedConnection)
        ? savedConnection
        : (connections[0] ? connections[0].name : '');
      return {
        consentRequired: false,
        environments: context.environments,
        environmentId: environment.id,
        connections,
        connectionName,
        flowId: id,
        numbering,
        flowMessage
      };
    } catch (error) {
      if (error instanceof FlowConsentError) {
        return { consentRequired: true, flowMessage: '' };
      }
      return { flowMessage: messageOf(error) };
    }
  }

  private async _save(announce: boolean): Promise<IConfigDraft | undefined> {
    const draft = this.state.draft;
    if (!draft) {
      return undefined;
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
      this.setState({ errorMessage: problem, statusMessage: '' });
      return undefined;
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
    const saved: IConfigDraft = { ...next, id };
    this._rule = saved;
    const running = this.state.numbering === 'Started';
    this.setState({
      saving: false,
      countEdited: false,
      draft: saved,
      statusMessage: announce
        ? (running
          ? 'Saved. The next item on this list receives the next number.'
          : 'Saved. Turn on numbering for this list.')
        : this.state.statusMessage
    });
    return saved;
  }

  private async _changeFlow(action: 'on' | 'off' | 'redeploy'): Promise<void> {
    let draft = this.state.draft;
    if (!draft) {
      return;
    }
    if (action !== 'off') {
      const saved = await this._save(false);
      if (!saved || saved.id === undefined) {
        return;
      }
      draft = saved;
    }
    if (draft.id === undefined) {
      this.setState({ flowMessage: 'Save the rule before turning numbering off.' });
      return;
    }
    this.setState({ working: true, flowMessage: '', errorMessage: '' });
    const environments = this.state.environments.length > 0
      ? this.state.environments
      : (await loadFlowContext(this.props.aadHttpClientFactory)).environments;
    const environment = environments.filter((item) => item.id === this.state.environmentId)[0] || environments[0];
    const connection = this.state.connectionName || (action === 'off' ? draft.flowConnection : '');
    if (action !== 'off' && !environment) {
      this.setState({ working: false, environments, flowMessage: 'Choose an environment.' });
      return;
    }
    if (action !== 'off' && !connection) {
      this.setState({ working: false, environments, flowMessage: 'Choose a SharePoint connection.' });
      return;
    }
    const savedEnvironment = draft.flowEnvironment;
    const sameEnvironment = !!environment && savedEnvironment === environment.id;
    const connectionChanged = (draft.flowConnection || '') !== connection;
    let flowId = draft.flowId || this.state.flowId;
    let numbering: FlowRunState = this.state.numbering;
    if (action === 'off') {
      if (flowId) {
        const stopEnvironment = environments.filter((item) => item.id === (savedEnvironment || (environment && environment.id)))[0];
        if (!stopEnvironment) {
          this.setState({ working: false, environments, flowMessage: 'The environment for this flow is not available to this account.' });
          return;
        }
        await setRemoteFlowState(this.props.aadHttpClientFactory, stopEnvironment, flowId, 'Stopped');
      }
      numbering = 'Stopped';
    } else if (environment) {
      const target = this._target();
      this._pending = { environment, connectionName: connection, itemId: draft.id };
      if (flowId && !sameEnvironment) {
        const previous = environments.filter((item) => item.id === savedEnvironment)[0];
        if (previous) {
          try {
            await setRemoteFlowState(this.props.aadHttpClientFactory, previous, flowId, 'Stopped');
          } catch (error) {
            if (error instanceof FlowConsentError) {
              throw error;
            }
          }
        }
        const remote = await createRemoteFlow(this.props.aadHttpClientFactory, environment, target, connection);
        flowId = remote.id;
        numbering = 'Started';
      } else if (flowId && (action === 'redeploy' || connectionChanged)) {
        const remote = await redeployRemoteFlow(this.props.aadHttpClientFactory, environment, flowId, target, connection);
        flowId = remote.id;
        numbering = 'Started';
      } else if (flowId && numbering === 'Stopped') {
        await setRemoteFlowState(this.props.aadHttpClientFactory, environment, flowId, 'Started');
        numbering = 'Started';
      } else if (!flowId) {
        const remote = await createRemoteFlow(this.props.aadHttpClientFactory, environment, target, connection);
        flowId = remote.id;
        numbering = 'Started';
      }
    }
    const stateName = numbering === 'Started' ? 'Started' : 'Stopped';
    const flowEnvironment = environment ? environment.id : savedEnvironment;
    await saveFlowBinding(this.props.spHttpClient, this.props.siteAbsoluteUrl, this._listTitle(), draft.id, {
      flowId,
      flowEnvironment,
      flowConnection: connection,
      flowState: stateName
    });
    const legacy = numbering === 'Started'
      ? await this._stopLegacySiteFlow(environments)
      : { note: '', legacy: this.state.legacySite };
    const updated: IConfigDraft = {
      ...draft,
      flowId,
      flowEnvironment,
      flowConnection: connection,
      flowState: stateName
    };
    this._rule = updated;
    this.setState({
      working: false,
      consentRequired: false,
      environments,
      environmentId: flowEnvironment,
      connectionName: connection,
      flowId,
      numbering,
      draft: updated,
      legacySite: legacy.legacy,
      flowMessage: legacy.note,
      statusMessage: numbering === 'Started'
        ? 'Numbering is on for this list.'
        : 'Numbering is off for this list. The saved rule stays in place.'
    });
  }

  private async _stopLegacySiteFlow(environments: IFlowEnvironment[]): Promise<{ note: string; legacy?: IConfigDraft }> {
    const legacy = this.state.legacySite;
    if (!legacy || !legacy.flowId || legacy.flowState === 'Stopped') {
      return { note: '', legacy };
    }
    const environment = environments.filter((item) => item.id === legacy.flowEnvironment)[0];
    if (!environment) {
      return {
        note: 'The previous site numbering flow is still on. Turn it off in Power Automate so each item is numbered once.',
        legacy
      };
    }
    try {
      await setRemoteFlowState(this.props.aadHttpClientFactory, environment, legacy.flowId, 'Stopped');
      const stopped: IConfigDraft = { ...legacy, flowState: 'Stopped' };
      await saveConfigRow(this.props.spHttpClient, this.props.siteAbsoluteUrl, this._listTitle(), stopped, false, true);
      return { note: '', legacy: stopped };
    } catch {
      return {
        note: 'The previous site numbering flow is still on. Turn it off in Power Automate so each item is numbered once.',
        legacy
      };
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

  private _failFlow(error: unknown): void {
    if (error instanceof FlowConsentError) {
      this.setState({ working: false, saving: false, consentRequired: true, flowMessage: '' });
      return;
    }
    if (error instanceof FlowStartError && this._pending && this._rule && this._rule.id !== undefined) {
      const pending = this._pending;
      const rule = this._rule;
      saveFlowBinding(this.props.spHttpClient, this.props.siteAbsoluteUrl, this._listTitle(), pending.itemId, {
        flowId: error.flowId,
        flowEnvironment: pending.environment.id,
        flowConnection: pending.connectionName,
        flowState: 'Stopped'
      }).then(() => {
        const draft: IConfigDraft = {
          ...rule,
          id: pending.itemId,
          flowId: error.flowId,
          flowEnvironment: pending.environment.id,
          flowConnection: pending.connectionName,
          flowState: 'Stopped'
        };
        this._rule = draft;
        this.setState({
          working: false,
          saving: false,
          draft,
          flowId: error.flowId,
          environmentId: pending.environment.id,
          connectionName: pending.connectionName,
          numbering: 'Stopped',
          flowMessage: error.message
        });
      }).catch(() => {
        this.setState({
          working: false,
          saving: false,
          flowId: error.flowId,
          numbering: 'Stopped',
          flowMessage: error.message
        });
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
