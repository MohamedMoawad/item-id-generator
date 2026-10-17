import * as React from 'react';
import * as ReactDom from 'react-dom';
import { BaseListViewCommandSet, IListViewCommandSetExecuteEventParameters } from '@microsoft/sp-listview-extensibility';
import { buildListAbsoluteUrl, DEFAULT_CONFIG_LIST_TITLE, isFlowConfigList } from '../../services/contract';
import FlowSettingsPanel from './FlowSettingsPanel';

export interface IFlowSettingsCommandSetProperties {
  configListTitle?: string;
}

export default class FlowSettingsCommandSet
  extends BaseListViewCommandSet<IFlowSettingsCommandSetProperties> {

  private _host: HTMLDivElement | undefined;

  public onInit(): Promise<void> {
    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);
    this._setCommandVisibility();
    return Promise.resolve();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    if (event.itemId !== 'AUTOGEN_FLOW') {
      return;
    }
    this._openPanel();
  }

  public onDispose(): void {
    this.context.listView.listViewStateChangedEvent.remove(this, this._onListViewStateChanged);
    this._closePanel();
  }

  private _onListViewStateChanged = (): void => {
    this._setCommandVisibility();
    this.raiseOnChange();
  };

  private _setCommandVisibility(): void {
    const command = this.tryGetCommand('AUTOGEN_FLOW');
    if (!command) {
      return;
    }
    const list = this.context.pageContext.list;
    const configTitle = this.properties.configListTitle || DEFAULT_CONFIG_LIST_TITLE;
    command.visible = !!list && !isFlowConfigList(list.title, list.serverRelativeUrl, configTitle);
    command.disabled = false;
    command.title = 'AutoGen Flow';
  }

  private _openPanel(): void {
    const list = this.context.pageContext.list;
    if (!list) {
      return;
    }
    if (!this._host) {
      this._host = document.body.appendChild(document.createElement('div'));
    }
    const element: React.ReactElement = React.createElement(FlowSettingsPanel, {
      siteAbsoluteUrl: this.context.pageContext.site.absoluteUrl,
      webAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
      configListTitle: this.properties.configListTitle || DEFAULT_CONFIG_LIST_TITLE,
      listTitle: list.title,
      listUrl: buildListAbsoluteUrl(this.context.pageContext.web.absoluteUrl, list.serverRelativeUrl),
      listGuid: list.id.toString(),
      spHttpClient: this.context.spHttpClient,
      aadHttpClientFactory: this.context.aadHttpClientFactory,
      siteTitle: this.context.pageContext.web.title,
      onDismiss: this._onDismiss
    });
    ReactDom.render(element, this._host);
  }

  private _onDismiss = (): void => {
    this._closePanel();
  };

  private _closePanel(): void {
    if (!this._host) {
      return;
    }
    ReactDom.unmountComponentAtNode(this._host);
  }
}
