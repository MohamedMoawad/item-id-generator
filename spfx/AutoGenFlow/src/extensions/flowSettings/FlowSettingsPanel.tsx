import * as React from 'react';
import {
  Checkbox,
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
  TextField
} from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import {
  DEFAULT_CONFIG_LIST_TITLE,
  draftForCurrentList,
  findConfigForList,
  IConfigDraft,
  RESET_PERIODS,
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
  showConfigList
} from '../../services/configListClient';
import { IFormulaColumn, planNextCode, prepareRule, textColumns } from '../../services/formula';
import {
  blankGuardExpression,
  configFilter,
  FLOW_CONCURRENCY,
  flowExpressions,
  flowSteps,
  triggerCondition
} from '../../services/flowGuide';
import styles from './FlowSettingsPanel.module.scss';

export interface IFlowSettingsPanelProps {
  siteAbsoluteUrl: string;
  webAbsoluteUrl: string;
  configListTitle: string;
  listTitle: string;
  listUrl: string;
  listGuid: string;
  spHttpClient: SPHttpClient;
  onDismiss: () => void;
}

interface IFlowSettingsPanelState {
  loading: boolean;
  saving: boolean;
  canEdit: boolean;
  countEdited: boolean;
  draft?: IConfigDraft;
  columns: IFormulaColumn[];
  errorMessage: string;
  statusMessage: string;
}

const DATE_TOKENS = ['{yyyy}', '{yy}', '{MM}', '{dd}', '{HH}', '{mm}', '{counter}', '{counter:4}'];

export default class FlowSettingsPanel
  extends React.Component<IFlowSettingsPanelProps, IFlowSettingsPanelState> {

  public constructor(props: IFlowSettingsPanelProps) {
    super(props);
    this.state = {
      loading: true,
      saving: false,
      canEdit: false,
      countEdited: false,
      columns: [],
      errorMessage: '',
      statusMessage: ''
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
        maxWidth={640}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'AutoGen Flow',
          showCloseButton: true
        }}
        modalProps={{
          isBlocking: false,
          containerClassName: styles.dialogHost
        }}
      >
        <div className={styles.panel}>
          <p className={styles.lead}>
            This saves the rule for {this.props.listTitle}. Power Automate writes the number when an item is created.
          </p>
          {this.state.loading && <Spinner size={SpinnerSize.small} label="Loading the master list" />}
          {this.state.errorMessage && <MessageBar messageBarType={MessageBarType.error}>{this.state.errorMessage}</MessageBar>}
          {this.state.statusMessage && <MessageBar messageBarType={MessageBarType.success}>{this.state.statusMessage}</MessageBar>}
          {draft && this._form(draft)}
        </div>
      </Dialog>
    );
  }

  private _form(draft: IConfigDraft): React.ReactElement {
    const numberOptions = this._numberOptions(draft);
    const prepared = prepareRule(draft.formula, draft.padLength);
    return (
      <>
        <Dropdown
          label="Number column"
          placeholder="Single line of text"
          options={numberOptions}
          selectedKey={draft.numberColumnInternalName || undefined}
          disabled={!this.state.canEdit || this.state.saving}
          onChange={this._onColumnChange}
        />
        <TextField
          label="Formula"
          description="Use {counter} or {counter:4}. Dates use the item Created time in UTC."
          multiline
          rows={3}
          value={draft.formula}
          disabled={!this.state.canEdit || this.state.saving}
          onChange={this._onFormulaChange}
        />
        <div className={styles.tokens}>
          {DATE_TOKENS.map((token) => (
            <button key={token} type="button" className={styles.token} data-token={token} onClick={this._onInsertToken}>{token}</button>
          ))}
          {this.state.columns.slice(0, 8).map((column) => (
            <button key={column.internalName} type="button" className={styles.token} data-token={column.token} onClick={this._onInsertToken}>
              {column.token}
            </button>
          ))}
        </div>
        <TextField
          label="Pad length"
          description="Used when the formula says {counter} without a width."
          value={String(draft.padLength)}
          disabled={!this.state.canEdit || this.state.saving}
          onChange={this._onPadChange}
        />
        <Dropdown
          label="Reset period"
          options={RESET_PERIODS.map((period) => ({ key: period, text: period }))}
          selectedKey={draft.resetPeriod}
          disabled={!this.state.canEdit || this.state.saving}
          onChange={this._onResetChange}
        />
        <TextField
          label="Current count"
          description="The last number used. The next item adds one. Set 0 to start at 1. Leave it unchanged and a save keeps the live counter."
          value={String(draft.currentCount)}
          disabled={!this.state.canEdit || this.state.saving}
          onChange={this._onCountChange}
        />
        <Checkbox
          label="Active"
          checked={draft.isActive}
          disabled={!this.state.canEdit || this.state.saving}
          onChange={this._onActiveChange}
        />
        <p className={styles.preview}>Next code from the saved counter: {this._preview(draft)}</p>
        <p className={styles.hint}>
          Stored for the flow: token {prepared.counterToken}, width {prepared.counterWidth}
          {prepared.formulaColumns ? `, columns ${prepared.formulaColumns}` : ''}.
        </p>
        <div className={styles.actions}>
          <PrimaryButton text="Save rule" disabled={!this.state.canEdit || this.state.saving || this.state.loading} onClick={this._save} />
          <DefaultButton text="Close" onClick={this.props.onDismiss} />
        </div>
        {draft.id !== undefined && (
          <details className={styles.flow} open>
            <summary>Power Automate for this list</summary>
            <p className={styles.step}>Master list site: {this.props.siteAbsoluteUrl}</p>
            <p className={styles.step}>Master list: {this.props.configListTitle || DEFAULT_CONFIG_LIST_TITLE}</p>
            <p className={styles.step}>Trigger list: {this.props.listUrl}</p>
            <p className={styles.step}>Concurrency: {FLOW_CONCURRENCY}</p>
            <TextField label="Filter query" readOnly multiline rows={2} value={configFilter(draft.targetListGuid)} />
            <TextField label="Trigger condition" readOnly multiline rows={2} value={triggerCondition(draft.numberColumnInternalName)} />
            <TextField label="Condition expression" readOnly multiline rows={3} value={blankGuardExpression(draft.numberColumnInternalName)} />
            {flowSteps().map((step, index) => (
              <p key={step} className={styles.step}>{index + 1}. {step}</p>
            ))}
            {flowExpressions(draft.numberColumnInternalName).map((item) => (
              <TextField
                key={item.action}
                className={styles.expression}
                label={item.action}
                readOnly
                multiline
                rows={3}
                value={item.expression}
              />
            ))}
          </details>
        )}
      </>
    );
  }

  private _numberOptions(draft: IConfigDraft): IDropdownOption[] {
    const options = textColumns(this.state.columns).map((column) => ({
      key: column.internalName,
      text: `${column.title} (${column.internalName})`
    }));
    if (draft.numberColumnInternalName && !options.some((option) => option.key === draft.numberColumnInternalName)) {
      options.unshift({
        key: draft.numberColumnInternalName,
        text: draft.numberColumnInternalName
      });
    }
    return options;
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
    this.setState({
      draft: { ...draft, formula: `${draft.formula}${token}` },
      statusMessage: ''
    });
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
    this.setState({
      draft: { ...draft, padLength: Number.isFinite(padLength) ? padLength : 0 },
      statusMessage: ''
    });
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

  private _onActiveChange = (_event?: React.FormEvent<HTMLElement | HTMLInputElement>, checked?: boolean): void => {
    const draft = this.state.draft;
    if (!draft) {
      return;
    }
    this.setState({ draft: { ...draft, isActive: checked === true }, statusMessage: '' });
  };

  private _save = (): void => {
    this._saveAsync().catch((error: unknown) => {
      this.setState({ saving: false, errorMessage: messageOf(error) });
    });
  };

  private async _load(): Promise<void> {
    const listTitle = this.props.configListTitle || DEFAULT_CONFIG_LIST_TITLE;
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
      errorMessage: canEdit ? '' : 'You can view this rule. Edit permission on the master list is required to save it.'
    });
  }

  private async _saveAsync(): Promise<void> {
    const draft = this.state.draft;
    if (!draft) {
      return;
    }
    const prepared = prepareRule(draft.formula, draft.padLength);
    const next: IConfigDraft = {
      ...draft,
      ...prepared,
      title: draft.title.trim() || this.props.listTitle,
      targetListTitle: this.props.listTitle,
      targetListUrl: this.props.listUrl,
      targetListGuid: this.props.listGuid
    };
    const problem = validateDraft(next);
    if (problem) {
      this.setState({ errorMessage: problem });
      return;
    }
    this.setState({ saving: true, errorMessage: '' });
    const listTitle = this.props.configListTitle || DEFAULT_CONFIG_LIST_TITLE;
    await ensureConfigList(this.props.spHttpClient, this.props.siteAbsoluteUrl, listTitle);
    const id = await saveConfigRow(
      this.props.spHttpClient,
      this.props.siteAbsoluteUrl,
      listTitle,
      next,
      next.id === undefined || this.state.countEdited
    );
    await ensureTargetNumberColumn(
      this.props.spHttpClient,
      this.props.listUrl,
      this.props.listGuid,
      next.numberColumnInternalName
    );
    this.setState({
      saving: false,
      countEdited: false,
      draft: { ...next, id },
      statusMessage: 'Saved. Create the Power Automate flow below. This page does not number items.'
    });
  }
}

function messageOf(error: unknown): string {
  if (error instanceof Error) {
    return error.message;
  }
  return 'The rule could not be saved.';
}
