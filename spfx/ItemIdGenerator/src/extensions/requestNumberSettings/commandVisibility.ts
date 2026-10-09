import { isConfigList } from '../../webparts/itemIdGenerator/services/configContract';

export function listServerRelativeUrlFromPath(pathname: string): string | undefined {
  const path = decodeURIComponent(pathname.split('?')[0]);
  const lists = path.match(/^(.*\/Lists\/[^/]+)/i);
  if (lists) {
    return lists[1];
  }
  const forms = path.match(/^(.*)\/Forms\/[^/]+\.aspx$/i);
  if (forms) {
    return forms[1];
  }
  return undefined;
}

export function shouldShowAutogenCommand(
  listTitle: string | undefined,
  serverRelativeUrl: string | undefined,
  configListTitle: string
): boolean {
  if (!listTitle || !serverRelativeUrl) {
    return true;
  }
  return !isConfigList(listTitle, serverRelativeUrl, configListTitle);
}
