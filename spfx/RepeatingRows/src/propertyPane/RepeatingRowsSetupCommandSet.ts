import { BaseListViewCommandSet, IListViewCommandSetExecuteEventParameters } from '@microsoft/sp-listview-extensibility';
import { SPHttpClient } from '@microsoft/sp-http';
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { SetupPanel } from '../controls/SetupPanel';
import { IRepeatingRowsConfig, RepeatingRowsError } from '../models/repeatingRows';
import { loadExistingConfig } from '../services/loadExistingConfig';
import { fieldHeaders, provisionRepeatingRows } from '../services/sharePointProvisioner';
import { SpHttpGateway } from '../services/sharePointGateway';

export interface IRepeatingRowsSetupCommandSetProperties {
}

const HOST_ID = 'repeating-rows-setup-host';

export default class RepeatingRowsSetupCommandSet extends BaseListViewCommandSet<IRepeatingRowsSetupCommandSetProperties> {
  private _session = 0;

  public onInit(): Promise<void> {
    const command = this.tryGetCommand('REPEATING_ROWS');
    if (command) {
      command.visible = !!this.context.pageContext.list;
    }
    return Promise.resolve();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    if (event.itemId !== 'REPEATING_ROWS') {
      return;
    }
    this._open();
  }

  public onDispose(): void {
    this._close();
  }

  private _open(): void {
    const list = this.context.pageContext.list;
    if (!list) {
      return;
    }
    const session = this._session + 1;
    this._session = session;
    const listId = list.id.toString();
    this._render(list.title, listId, undefined, false);
    const gateway = new SpHttpGateway(this.context.spHttpClient, this.context.pageContext.web.absoluteUrl);
    loadExistingConfig(gateway, listId, `${list.title} Details`).then((existing) => {
      if (session === this._session) {
        this._render(list.title, listId, existing, true);
      }
    }).catch(() => {
      if (session === this._session) {
        this._render(list.title, listId, undefined, true);
      }
    });
  }

  private _render(listTitle: string, listId: string, existing: IRepeatingRowsConfig | undefined, ready: boolean): void {
    const host = document.getElementById(HOST_ID) || document.createElement('div');
    host.id = HOST_ID;
    if (!host.parentElement) {
      document.body.appendChild(host);
    }
    ReactDOM.render(React.createElement(SetupPanel, {
      listTitle,
      ready,
      existing,
      onDismiss: () => this._close(),
      onProvision: (config) => this._provision(listId, config)
    }), host);
  }

  private _close(): void {
    this._session += 1;
    const host = document.getElementById(HOST_ID);
    if (!host) {
      return;
    }
    ReactDOM.unmountComponentAtNode(host);
    host.remove();
  }

  private _provision(listId: string, config: IRepeatingRowsConfig): Promise<void> {
    const gateway = new SpHttpGateway(this.context.spHttpClient, this.context.pageContext.web.absoluteUrl);
    return provisionRepeatingRows(gateway, listId, config, (url, body) => postField(this.context.spHttpClient, url, body));
  }
}

async function postField(client: SPHttpClient, url: string, body: string): Promise<{ status: number; text: string }> {
  const response = await client.post(url, SPHttpClient.configurations.v1, {
    headers: fieldHeaders,
    body
  });
  return { status: response.status, text: await response.text() };
}

export function requireList(listId: string | undefined): string {
  if (!listId) {
    throw new RepeatingRowsError('Open a list before configuring repeating rows.');
  }
  return listId;
}
