import { flowListUrl } from './flowModel';

interface IStorageDocument extends Document {
  requestStorageAccessFor?: (origin: string) => Promise<void>;
}

export function beginPowerAutomateAccess(): void {
  const request = (document as IStorageDocument).requestStorageAccessFor;
  if (!request) {
    return;
  }
  try {
    request.call(document, 'https://make.powerautomate.com').catch(() => undefined);
  } catch {
    return;
  }
}

export function openPowerAutomateSignIn(environmentId: string): Window | undefined {
  const url = environmentId.trim() ? flowListUrl(environmentId) : 'https://make.powerautomate.com/flows';
  return window.open(url, '_blank') || undefined;
}

export function showPowerAutomatePage(tab: Window | undefined, url: string): void {
  if (tab && !tab.closed) {
    try {
      tab.location.href = url;
      return;
    } catch {
      window.open(url, '_blank', 'noopener,noreferrer');
      return;
    }
  }
  window.open(url, '_blank', 'noopener,noreferrer');
}
