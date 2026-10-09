import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { SPHttpClient } from '@microsoft/sp-http';
import { buildGetListUrl, DEFAULT_CONFIG_LIST_TITLE } from '../../webparts/itemIdGenerator/services/configContract';
import { ensureConfigList, hideConfigList } from '../../webparts/itemIdGenerator/services/configListClient';
import { syncAutogenFromCustomizer } from '../requestNumberSettings/autogenListRuntime';
import { listServerRelativeUrlFromPath } from '../requestNumberSettings/commandVisibility';

export interface IRequestNumberProvisionerProperties {
  configListTitle?: string;
}

const ENSURED_KEY = 'requestNumberConfigEnsured';

export default class RequestNumberProvisionerApplicationCustomizer
  extends BaseApplicationCustomizer<IRequestNumberProvisionerProperties> {

  private _timer: number | undefined;
  private _resolvedPath = '';
  private _resolving = false;

  public onInit(): Promise<void> {
    this._syncList();
    this._timer = window.setInterval(() => this._syncList(), 1000);
    return this._ensureConfigList();
  }

  public onDispose(): void {
    if (this._timer !== undefined) {
      window.clearInterval(this._timer);
    }
    syncAutogenFromCustomizer(undefined);
  }

  private _syncList(): void {
    const list = this.context.pageContext.list;
    const properties = this.properties || {};
    const configListTitle = properties.configListTitle || DEFAULT_CONFIG_LIST_TITLE;
    if (list) {
      this._resolvedPath = list.serverRelativeUrl;
      syncAutogenFromCustomizer({
        spHttpClient: this.context.spHttpClient,
        siteAbsoluteUrl: this.context.pageContext.site.absoluteUrl,
        webAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
        configListTitle,
        listId: list.id.toString(),
        listTitle: list.title,
        listServerRelativeUrl: list.serverRelativeUrl
      });
      return;
    }
    const serverRelative = listServerRelativeUrlFromPath(window.location.pathname);
    if (!serverRelative) {
      this._resolvedPath = '';
      syncAutogenFromCustomizer(undefined);
      return;
    }
    if (this._resolvedPath === serverRelative || this._resolving) {
      return;
    }
    this._resolving = true;
    this._resolveList(serverRelative, configListTitle).catch(() => undefined);
  }

  private async _resolveList(serverRelative: string, configListTitle: string): Promise<void> {
    try {
      const origin = new URL(this.context.pageContext.web.absoluteUrl).origin;
      const response = await this.context.spHttpClient.get(
        buildGetListUrl(`${origin}${serverRelative}`),
        SPHttpClient.configurations.v1,
        { headers: { Accept: 'application/json;odata=nometadata' } }
      );
      if (!response.ok) {
        return;
      }
      const body = await response.json() as { Id?: string; Title?: string };
      if (!body.Id || !body.Title) {
        return;
      }
      this._resolvedPath = serverRelative;
      syncAutogenFromCustomizer({
        spHttpClient: this.context.spHttpClient,
        siteAbsoluteUrl: this.context.pageContext.site.absoluteUrl,
        webAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
        configListTitle,
        listId: body.Id,
        listTitle: body.Title,
        listServerRelativeUrl: serverRelative
      });
    } finally {
      this._resolving = false;
    }
  }

  private async _ensureConfigList(): Promise<void> {
    try {
      if (window.sessionStorage.getItem(ENSURED_KEY) === '1') {
        return;
      }
    } catch {
      // Session storage can be blocked. Still try to create the list.
    }

    const listTitle = (this.properties.configListTitle || DEFAULT_CONFIG_LIST_TITLE).trim() || DEFAULT_CONFIG_LIST_TITLE;
    try {
      await ensureConfigList(
        this.context.spHttpClient,
        this.context.pageContext.site.absoluteUrl,
        listTitle
      );
      await hideConfigList(
        this.context.spHttpClient,
        this.context.pageContext.site.absoluteUrl,
        listTitle
      );
      window.sessionStorage.setItem(ENSURED_KEY, '1');
    } catch {
      // A reader may not be allowed to create the list. The next owner visit retries.
    }
  }
}
