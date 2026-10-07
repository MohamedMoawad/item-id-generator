import * as React from 'react';
import * as ReactDom from 'react-dom';
import { BaseListViewCommandSet, IListViewCommandSetExecuteEventParameters } from '@microsoft/sp-listview-extensibility';
import { DEFAULT_CONFIG_LIST_TITLE, buildListAbsoluteUrl, isConfigList } from '../../webparts/itemIdGenerator/services/configContract';
import RequestNumberSettingsPanel from './RequestNumberSettingsPanel';

export interface IRequestNumberSettingsCommandSetProperties {
  configListTitle?: string;
}

export default class RequestNumberSettingsCommandSet
  extends BaseListViewCommandSet<IRequestNumberSettingsCommandSetProperties> {

  private _panelHost: HTMLDivElement | undefined;

  public onInit(): Promise<void> {
    this._panelHost = document.body.appendChild(document.createElement('div'));
    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);
    this._setCommandVisibility();
    return Promise.resolve();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    if (event.itemId !== 'REQUEST_NUMBER_SETTINGS') {
      return;
    }
    this._openPanel();
  }

  public onDispose(): void {
    this.context.listView.listViewStateChangedEvent.remove(this, this._onListViewStateChanged);
    this._closePanel();
    if (this._panelHost && this._panelHost.parentElement) {
      this._panelHost.parentElement.removeChild(this._panelHost);
    }
  }

  private _onListViewStateChanged = (): void => {
    this._setCommandVisibility();
    this.raiseOnChange();
  };

  private _setCommandVisibility(): void {
    const command = this.tryGetCommand('REQUEST_NUMBER_SETTINGS');
    if (!command) {
      return;
    }
    const list = this.context.pageContext.list;
    const title = list ? list.title : '';
    const url = list ? list.serverRelativeUrl : '';
    const configTitle = this.properties.configListTitle || DEFAULT_CONFIG_LIST_TITLE;
    command.visible = !!list && !isConfigList(title, url, configTitle);
  }

  private _openPanel(): void {
    const list = this.context.pageContext.list;
    if (!list || !this._panelHost) {
      return;
    }
    const element: React.ReactElement = React.createElement(RequestNumberSettingsPanel, {
      siteAbsoluteUrl: this.context.pageContext.site.absoluteUrl,
      configListTitle: this.properties.configListTitle || DEFAULT_CONFIG_LIST_TITLE,
      listTitle: list.title,
      listUrl: buildListAbsoluteUrl(this.context.pageContext.web.absoluteUrl, list.serverRelativeUrl),
      listGuid: list.id.toString(),
      spHttpClient: this.context.spHttpClient,
      onDismiss: this._closePanel
    });
    ReactDom.render(element, this._panelHost);
  }

  private _closePanel = (): void => {
    if (this._panelHost) {
      ReactDom.unmountComponentAtNode(this._panelHost);
    }
  };
}
