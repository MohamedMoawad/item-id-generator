import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { DEFAULT_CONFIG_LIST_TITLE, isFlowConfigList } from '../../services/contract';
import { ensureConfigList, showConfigList } from '../../services/configListClient';
import { watchListNumbers } from '../../services/listFill';

export interface IFlowProvisionerProperties {
  configListTitle?: string;
}

const ENSURED_KEY = 'autogenFlowConfigVisible';

export default class FlowProvisionerApplicationCustomizer
  extends BaseApplicationCustomizer<IFlowProvisionerProperties> {

  public onInit(): Promise<void> {
    this.context.application.navigatedEvent.add(this, this._onNavigated);
    this._watchCurrentList();
    return this._ensureConfigList();
  }

  public onDispose(): void {
    this.context.application.navigatedEvent.remove(this, this._onNavigated);
    watchListNumbers(undefined);
  }

  private _onNavigated = (): void => {
    this._watchCurrentList();
  };

  private _watchCurrentList(): void {
    const list = this.context.pageContext.list;
    const listTitle = this._configListTitle();
    if (!list || isFlowConfigList(list.title, list.serverRelativeUrl, listTitle)) {
      watchListNumbers(undefined);
      return;
    }
    watchListNumbers({
      spHttpClient: this.context.spHttpClient,
      siteAbsoluteUrl: this.context.pageContext.site.absoluteUrl,
      webAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
      configListTitle: listTitle,
      listGuid: list.id.toString()
    });
  }

  private async _ensureConfigList(): Promise<void> {
    try {
      if (window.sessionStorage.getItem(ENSURED_KEY) === '1') {
        return;
      }
    } catch {
      // Session storage can be blocked. Still try to create the list.
    }

    const listTitle = this._configListTitle();
    try {
      await ensureConfigList(
        this.context.spHttpClient,
        this.context.pageContext.site.absoluteUrl,
        listTitle
      );
      await showConfigList(
        this.context.spHttpClient,
        this.context.pageContext.site.absoluteUrl,
        listTitle
      );
      window.sessionStorage.setItem(ENSURED_KEY, '1');
    } catch {
      // A reader may not be allowed to create the list. The next owner visit retries.
    }
  }

  private _configListTitle(): string {
    const properties = this.properties || {};
    return (properties.configListTitle || DEFAULT_CONFIG_LIST_TITLE).trim() || DEFAULT_CONFIG_LIST_TITLE;
  }
}
