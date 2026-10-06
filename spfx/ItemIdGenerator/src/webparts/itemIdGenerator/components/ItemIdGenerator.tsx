import * as React from 'react';
import { MessageBar, MessageBarType, PrimaryButton, Spinner, SpinnerSize, TextField } from '@fluentui/react';
import styles from './ItemIdGenerator.module.scss';
import type { IItemIdGeneratorProps } from './IItemIdGeneratorProps';
import { requestGeneratedItemId } from '../services/generateItemIdClient';
import { updateListItemField } from '../services/updateListItemField';

interface IItemIdGeneratorState {
  itemIdText: string;
  busy: boolean;
  generatedId: string;
  errorMessage: string;
  savedMessage: string;
}

export default class ItemIdGenerator extends React.Component<IItemIdGeneratorProps, IItemIdGeneratorState> {
  constructor(props: IItemIdGeneratorProps) {
    super(props);
    this.state = {
      itemIdText: '',
      busy: false,
      generatedId: '',
      errorMessage: '',
      savedMessage: ''
    };
  }

  public render(): React.ReactElement<IItemIdGeneratorProps> {
    const missingSettings = this._missingSettings();
    const conditionLabel = this._conditionLabel();

    return (
      <section className={styles.itemIdGenerator}>
        <h2 className={styles.header}>Item ID</h2>
        <p className={styles.lede}>
          Generate a conditional ID for one list item, then write it to the target column.
          Default IDs look like Requests-timestamp. VIP IDs are prefixed with VIP-.
        </p>
        <ul className={styles.settings}>
          <li>List: {this.props.listName.trim() || 'Not set'}</li>
          <li>Field: {this.props.targetFieldInternalName.trim() || 'Not set'}</li>
          <li>Condition: {conditionLabel}</li>
        </ul>
        {missingSettings.length > 0 && (
          <div className={styles.status}>
            <MessageBar messageBarType={MessageBarType.warning}>
              Set {missingSettings.join(', ')} in the property pane before generating an ID.
            </MessageBar>
          </div>
        )}
        <div className={styles.form}>
          <TextField
            label="List item ID"
            description="Numeric ID of the item that should receive the generated value."
            value={this.state.itemIdText}
            onChange={this._onItemIdChange}
            disabled={this.state.busy}
            required={true}
          />
          <div className={styles.actions}>
            <PrimaryButton
              text={this.state.busy ? 'Generating...' : 'Generate ID'}
              onClick={this._onGenerateClick}
              disabled={this.state.busy || missingSettings.length > 0}
            />
          </div>
        </div>
        {this.state.busy && (
          <div className={styles.status}>
            <Spinner size={SpinnerSize.small} label="Generating an ID and updating the list item" />
          </div>
        )}
        {this.state.errorMessage && (
          <div className={styles.status}>
            <MessageBar messageBarType={MessageBarType.error}>{this.state.errorMessage}</MessageBar>
          </div>
        )}
        {this.state.savedMessage && (
          <div className={styles.status}>
            <MessageBar messageBarType={MessageBarType.success}>{this.state.savedMessage}</MessageBar>
          </div>
        )}
        {this.state.generatedId && (
          <div aria-live="polite">
            <div className={styles.resultLabel}>Generated ID</div>
            <p className={styles.resultValue}>{this.state.generatedId}</p>
          </div>
        )}
      </section>
    );
  }

  private _onItemIdChange = (
    _event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ): void => {
    this.setState({ itemIdText: newValue ?? '' });
  };

  private _onGenerateClick = (): void => {
    this._generateId().catch((error: unknown) => {
      const message = error instanceof Error ? error.message : 'ID generation failed.';
      this.setState({ busy: false, errorMessage: message });
    });
  };

  private async _generateId(): Promise<void> {
    const itemIdText = this.state.itemIdText.trim();
    if (!/^[1-9]\d*$/.test(itemIdText)) {
      this.setState({
        errorMessage: 'Enter the SharePoint list item ID (a positive integer).',
        savedMessage: '',
        generatedId: ''
      });
      return;
    }

    const itemId = Number(itemIdText);
    this.setState({
      busy: true,
      errorMessage: '',
      savedMessage: '',
      generatedId: ''
    });

    try {
      const generatedId = await requestGeneratedItemId({
        functionUrl: this.props.functionUrl,
        listName: this.props.listName,
        condition: this.props.condition || 'default',
        itemId
      });
      this.setState({ generatedId });

      try {
        await updateListItemField(this.props.spHttpClient, {
          webAbsoluteUrl: this.props.webAbsoluteUrl,
          listName: this.props.listName,
          itemId,
          targetFieldInternalName: this.props.targetFieldInternalName,
          generatedId
        });
        this.setState({
          savedMessage: `Saved to ${this.props.listName.trim()} item ${itemId}, field ${this.props.targetFieldInternalName.trim()}.`
        });
      } catch (updateError) {
        const message = updateError instanceof Error
          ? updateError.message
          : 'The list item update failed.';
        this.setState({ errorMessage: message });
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : 'ID generation failed.';
      this.setState({ errorMessage: message, generatedId: '' });
    } finally {
      this.setState({ busy: false });
    }
  }

  private _missingSettings(): string[] {
    const missing: string[] = [];
    const functionUrl = this.props.functionUrl.trim();
    if (!functionUrl || functionUrl.indexOf('<') >= 0) {
      missing.push('Function URL');
    }
    if (!this.props.listName.trim()) {
      missing.push('List name');
    }
    if (!this.props.targetFieldInternalName.trim()) {
      missing.push('Target field internal name');
    }
    return missing;
  }

  private _conditionLabel(): string {
    if (this.props.condition === 'VIP') {
      return 'VIP';
    }
    if (!this.props.condition || this.props.condition === 'default') {
      return 'Default';
    }
    return this.props.condition;
  }
}
