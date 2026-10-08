import * as React from 'react';
import * as ReactDom from 'react-dom';
import { BaseListViewCommandSet, IListViewCommandSetExecuteEventParameters } from '@microsoft/sp-listview-extensibility';
import { DEFAULT_CONFIG_LIST_TITLE, buildListAbsoluteUrl, isConfigList } from '../../webparts/itemIdGenerator/services/configContract';
import { connectConfiguredList } from '../../webparts/itemIdGenerator/services/configListClient';
import { numberBlankItems } from '../../webparts/itemIdGenerator/services/listNumbering';
import RequestNumberSettingsPanel from './RequestNumberSettingsPanel';

export interface IRequestNumberSettingsCommandSetProperties {
  configListTitle?: string;
}

export default class RequestNumberSettingsCommandSet
  extends BaseListViewCommandSet<IRequestNumberSettingsCommandSetProperties> {

  private _panelHost: HTMLDivElement | undefined;
  private _numbering = false;
  private _timer: number | undefined;

  public onInit(): Promise<void> {
    this._panelHost = document.body.appendChild(document.createElement('div'));
    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);
    this._setCommandVisibility();
    this._connectList().catch(() => undefined);
    this._numberList().catch(() => undefined);
    this._timer = window.setInterval(() => {
      this._numberList().catch(() => undefined);
    }, 8000);
    return Promise.resolve();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    if (event.itemId !== 'REQUEST_NUMBER_SETTINGS') {
      return;
    }
    this._openPanel();
  }

  public onDispose(): void {
    if (this._timer !== undefined) {
      window.clearInterval(this._timer);
    }
    this.context.listView.listViewStateChangedEvent.remove(this, this._onListViewStateChanged);
    this._closePanel();
    if (this._panelHost && this._panelHost.parentElement) {
      this._panelHost.parentElement.removeChild(this._panelHost);
    }
  }

  private _onListViewStateChanged = (): void => {
    this._setCommandVisibility();
    this._numberList().catch(() => undefined);
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

  private async _numberList(): Promise<void> {
    if (this._numbering) {
      return;
    }
    const list = this.context.pageContext.list;
    if (!list) {
      return;
    }
    const configTitle = this.properties.configListTitle || DEFAULT_CONFIG_LIST_TITLE;
    if (isConfigList(list.title, list.serverRelativeUrl, configTitle)) {
      return;
    }
    this._numbering = true;
    try {
      const written = await numberBlankItems(
        this.context.spHttpClient,
        this.context.pageContext.site.absoluteUrl,
        configTitle,
        this.context.pageContext.web.absoluteUrl,
        list.id.toString()
      );
      if (written > 0) {
        refreshListView();
      }
    } finally {
      this._numbering = false;
    }
  }

  private async _connectList(): Promise<void> {
    const list = this.context.pageContext.list;
    if (!list) {
      return;
    }
    const configTitle = this.properties.configListTitle || DEFAULT_CONFIG_LIST_TITLE;
    if (isConfigList(list.title, list.serverRelativeUrl, configTitle)) {
      return;
    }
    const key = `requestNumberConnected:${list.id.toString()}`;
    try {
      if (window.sessionStorage.getItem(key) === '1') {
        return;
      }
    } catch {
      // Session storage can be blocked. Still try to connect.
    }
    const connected = await connectConfiguredList(
      this.context.spHttpClient,
      this.context.pageContext.site.absoluteUrl,
      configTitle,
      list.id.toString()
    );
    if (!connected) {
      return;
    }
    try {
      window.sessionStorage.setItem(key, '1');
    } catch {
      // The connection itself succeeded.
    }
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

function refreshListView(): void {
  const button = document.querySelector('[data-automationid="refreshCommand"], button[name="Refresh"]');
  if (button instanceof HTMLButtonElement) {
    button.click();
  }
}
