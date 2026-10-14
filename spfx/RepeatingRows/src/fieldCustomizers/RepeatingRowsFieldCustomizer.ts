import { BaseFieldCustomizer, IFieldCustomizerCellEventParameters } from '@microsoft/sp-listview-extensibility';
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { FieldSummary } from '../controls/FieldSummary';
import { summaryText } from '../utilities/summary';

export interface IRepeatingRowsFieldProperties {
  childListTitle?: string;
}

export default class RepeatingRowsFieldCustomizer extends BaseFieldCustomizer<IRepeatingRowsFieldProperties> {
  public onInit(): Promise<void> {
    return Promise.resolve();
  }

  public onRenderCell(event: IFieldCustomizerCellEventParameters): void {
    ReactDOM.render(
      React.createElement(FieldSummary, { text: summaryText(event.fieldValue) }),
      event.domElement
    );
  }

  public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
    ReactDOM.unmountComponentAtNode(event.domElement);
  }
}
