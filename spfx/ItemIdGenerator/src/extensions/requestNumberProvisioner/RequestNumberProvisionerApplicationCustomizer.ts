import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { DEFAULT_CONFIG_LIST_TITLE } from '../../webparts/itemIdGenerator/services/configContract';
import { ensureConfigList } from '../../webparts/itemIdGenerator/services/configListClient';

export interface IRequestNumberProvisionerProperties {
  configListTitle?: string;
}

const ENSURED_KEY = 'requestNumberConfigEnsured';

export default class RequestNumberProvisionerApplicationCustomizer
  extends BaseApplicationCustomizer<IRequestNumberProvisionerProperties> {

  public onInit(): Promise<void> {
    return this._ensureConfigList();
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
      window.sessionStorage.setItem(ENSURED_KEY, '1');
    } catch {
      // A reader may not be allowed to create the list. The next owner visit retries.
    }
  }
}
