import { BaseListViewCommandSet, IListViewCommandSetExecuteEventParameters } from '@microsoft/sp-listview-extensibility';
import { SPHttpClient } from '@microsoft/sp-http';
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { SetupPanel } from '../controls/SetupPanel';
import { IRepeatingRowsConfig, RepeatingRowsError } from '../models/repeatingRows';
import { fieldHeaders, provisionRepeatingRows } from '../services/sharePointProvisioner';
import { SpHttpGateway } from '../services/sharePointGateway';

export interface IRepeatingRowsSetupCommandSetProperties {
}

const HOST_ID = 'repeating-rows-setup-host';

export default class RepeatingRowsSetupCommandSet extends BaseListViewCommandSet<IRepeatingRowsSetupCommandSetProperties> {
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
    const host = document.createElement('div');
    host.id = HOST_ID;
    document.body.appendChild(host);
    const element = React.createElement(SetupPanel, {
      listTitle: list.title,
      onDismiss: () => this._close(),
      onProvision: (config) => this._provision(list.id.toString(), config)
    });
    ReactDOM.render(element, host);
  }

  private _close(): void {
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
