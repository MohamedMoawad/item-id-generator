import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { DEFAULT_CONFIG_LIST_TITLE, isAutoGenSystemList } from '../../services/contract';
import { ensureConfigList, showConfigList } from '../../services/configListClient';
import { ensureLogList, showLogList } from '../../services/flowLog';
import { watchListNumbers } from '../../services/listFill';

export interface IFlowProvisionerProperties {
  configListTitle?: string;
}

const ENSURED_KEY = 'autogenFlowConfigVisible';
const LOG_KEY = 'autogenFlowLogVisible';

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
    if (!list || isAutoGenSystemList(list.title, list.serverRelativeUrl, listTitle)) {
      watchListNumbers(undefined);
      return;
    }
    watchListNumbers({
      spHttpClient: this.context.spHttpClient,
      siteAbsoluteUrl: this.context.pageContext.site.absoluteUrl,
      webAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
      configListTitle: listTitle,
      listGuid: list.id.toString(),
      listTitle: list.title
    });
  }

  private async _ensureConfigList(): Promise<void> {
    let configReady = false;
    try {
      configReady = window.sessionStorage.getItem(ENSURED_KEY) === '1';
    } catch {
      configReady = false;
    }
    if (!configReady) {
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
    await this._ensureLogList();
  }

  private async _ensureLogList(): Promise<void> {
    try {
      if (window.sessionStorage.getItem(LOG_KEY) === '1') {
        return;
      }
    } catch {
      // Session storage can be blocked. Still try to create the log list.
    }
    const site = this.context.pageContext.site.absoluteUrl;
    try {
      await ensureLogList(this.context.spHttpClient, site);
      await showLogList(this.context.spHttpClient, site);
      window.sessionStorage.setItem(LOG_KEY, '1');
    } catch {
      // A reader may not be allowed to create the log list. The next owner visit retries.
    }
  }

  private _configListTitle(): string {
    const properties = this.properties || {};
    return (properties.configListTitle || DEFAULT_CONFIG_LIST_TITLE).trim() || DEFAULT_CONFIG_LIST_TITLE;
  }
}
