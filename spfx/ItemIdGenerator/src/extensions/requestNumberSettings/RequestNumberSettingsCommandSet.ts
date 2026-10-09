import { BaseListViewCommandSet, IListViewCommandSetExecuteEventParameters } from '@microsoft/sp-listview-extensibility';
import { DEFAULT_CONFIG_LIST_TITLE } from '../../webparts/itemIdGenerator/services/configContract';
import { AUTOGEN_COMMAND_LABEL, openAutogenSettingsPanel, syncAutogenFromCommandSet } from './autogenListRuntime';
import { shouldShowAutogenCommand } from './commandVisibility';

export interface IRequestNumberSettingsCommandSetProperties {
  configListTitle?: string;
}

const AUTOGEN_ICON = 'data:image/svg+xml,' + encodeURIComponent(
  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 16 16"><rect width="16" height="16" rx="2" fill="#038387"/><text x="8" y="12" text-anchor="middle" font-size="11" font-family="Segoe UI" fill="#ffffff">#</text></svg>'
);

export default class RequestNumberSettingsCommandSet
  extends BaseListViewCommandSet<IRequestNumberSettingsCommandSetProperties> {

  public onInit(): Promise<void> {
    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);
    this._setCommandVisibility();
    this._syncRuntime();
    this.raiseOnChange();
    return Promise.resolve();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    if (event.itemId !== 'REQUEST_NUMBER_SETTINGS') {
      return;
    }
    this._syncRuntime();
    openAutogenSettingsPanel();
  }

  public onDispose(): void {
    this.context.listView.listViewStateChangedEvent.remove(this, this._onListViewStateChanged);
    syncAutogenFromCommandSet(undefined);
  }

  private _onListViewStateChanged = (): void => {
    this._setCommandVisibility();
    this._syncRuntime();
    this.raiseOnChange();
  };

  private _setCommandVisibility(): void {
    const command = this.tryGetCommand('REQUEST_NUMBER_SETTINGS');
    if (!command) {
      return;
    }
    const list = this.context.pageContext.list;
    const configTitle = this.properties.configListTitle || DEFAULT_CONFIG_LIST_TITLE;
    command.visible = shouldShowAutogenCommand(
      list ? list.title : undefined,
      list ? list.serverRelativeUrl : undefined,
      configTitle
    );
    command.disabled = false;
    command.title = AUTOGEN_COMMAND_LABEL;
    if (!command.iconImageUrl) {
      command.iconImageUrl = AUTOGEN_ICON;
    }
  }

  private _syncRuntime(): void {
    const list = this.context.pageContext.list;
    if (!list) {
      return;
    }
    syncAutogenFromCommandSet({
      spHttpClient: this.context.spHttpClient,
      siteAbsoluteUrl: this.context.pageContext.site.absoluteUrl,
      webAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
      configListTitle: this.properties.configListTitle || DEFAULT_CONFIG_LIST_TITLE,
      listId: list.id.toString(),
      listTitle: list.title,
      listServerRelativeUrl: list.serverRelativeUrl
    });
  }
}
