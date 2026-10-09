import { isConfigList } from '../../webparts/itemIdGenerator/services/configContract';

export interface ILegacyPageContextInfo {
  listId?: string;
  listTitle?: string;
  listUrl?: string;
}

export interface ILegacyListContext {
  listId: string;
  listTitle: string;
  listServerRelativeUrl: string;
}

export function readLegacyListContext(info: ILegacyPageContextInfo | undefined): ILegacyListContext | undefined {
  if (!info || !info.listId || !info.listTitle || !info.listUrl) {
    return undefined;
  }
  const listId = info.listId.replace(/[{}]/g, '').toLowerCase();
  if (!listId || listId === '00000000-0000-0000-0000-000000000000') {
    return undefined;
  }
  return {
    listId,
    listTitle: info.listTitle,
    listServerRelativeUrl: serverRelativePath(info.listUrl)
  };
}

function serverRelativePath(listUrl: string): string {
  const bare = listUrl.split('?')[0];
  if (bare.indexOf('https://') === 0 || bare.indexOf('http://') === 0) {
    return new URL(bare).pathname;
  }
  return bare.charAt(0) === '/' ? bare : `/${bare}`;
}

export function listServerRelativeUrlFromPath(pathname: string): string | undefined {
  let path = pathname.split('?')[0];
  try {
    path = decodeURIComponent(path);
  } catch {
    return undefined;
  }
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
