const WATCH_FLAG = '__autoGenListCreateWatch';

interface IWatchWindow extends Window {
  __autoGenListCreateWatch?: boolean;
}

export function isNewItemRequest(url: string, method: string, listGuid: string, serverRelativeUrl: string): boolean {
  if ((method || '').toUpperCase() !== 'POST') {
    return false;
  }
  let text = url;
  try {
    text = decodeURIComponent(url);
  } catch {
    text = url;
  }
  const lower = text.toLowerCase();
  const onThisList = lower.indexOf(listGuid.toLowerCase()) >= 0 ||
    (serverRelativeUrl.length > 1 && lower.indexOf(serverRelativeUrl.toLowerCase()) >= 0);
  if (!onThisList) {
    return false;
  }
  if (lower.indexOf('addvalidateupdateitemusingpath') >= 0) {
    return true;
  }
  return /\/items(\?|$)/i.test(lower) && !/\/items\(\d+\)/i.test(lower);
}

export function createdItemId(payload: unknown): number | undefined {
  const record = asRecord(payload);
  const direct = positiveId(record.Id) || positiveId(asRecord(record.d).Id);
  if (direct) {
    return direct;
  }
  const nested = asRecord(asRecord(record.d).AddValidateUpdateItemUsingPath);
  const rows = arrayOf(record.value) || arrayOf(nested.results) || arrayOf(nested.value);
  for (let index = 0; index < rows.length; index += 1) {
    const row = asRecord(rows[index]);
    const name = typeof row.FieldName === 'string' ? row.FieldName : '';
    if (name === 'Id' || name === 'ID') {
      const id = positiveId(row.FieldValue);
      if (id) {
        return id;
      }
    }
  }
  return undefined;
}

export function watchNewListItems(
  listGuid: string,
  serverRelativeUrl: string,
  onCreated: () => Promise<void>
): () => void {
  const target = window as IWatchWindow;
  if (target[WATCH_FLAG]) {
    return () => undefined;
  }
  target[WATCH_FLAG] = true;
  const originalFetch = window.fetch.bind(window);
  const originalOpen = XMLHttpRequest.prototype.open;
  const originalSend = XMLHttpRequest.prototype.send;

  window.fetch = (input: RequestInfo | URL, init?: RequestInit): Promise<Response> => {
    const url = requestUrl(input);
    const method = (init && init.method) || (input instanceof Request ? input.method : 'GET');
    return originalFetch(input, init).then(async (response) => {
      if (response.ok && isNewItemRequest(url, method, listGuid, serverRelativeUrl)) {
        await notifyIfCreated(response.clone(), onCreated);
      }
      return response;
    });
  };

  XMLHttpRequest.prototype.open = function (this: XMLHttpRequest, method: string, url: string | URL): void {
    (this as XMLHttpRequest & { __autoGenMethod?: string; __autoGenUrl?: string }).__autoGenMethod = method;
    (this as XMLHttpRequest & { __autoGenUrl?: string }).__autoGenUrl = String(url);
    return originalOpen.apply(this, arguments as unknown as Parameters<typeof originalOpen>);
  };

  XMLHttpRequest.prototype.send = function (this: XMLHttpRequest): void {
    const marked = this as XMLHttpRequest & { __autoGenMethod?: string; __autoGenUrl?: string };
    this.addEventListener('load', () => {
      const url = marked.__autoGenUrl || this.responseURL || '';
      const method = marked.__autoGenMethod || 'GET';
      if (this.status < 200 || this.status >= 300 || !isNewItemRequest(url, method, listGuid, serverRelativeUrl)) {
        return;
      }
      onCreated().catch(() => undefined);
    });
    return originalSend.apply(this, arguments as unknown as Parameters<typeof originalSend>);
  };

  return () => {
    window.fetch = originalFetch;
    XMLHttpRequest.prototype.open = originalOpen;
    XMLHttpRequest.prototype.send = originalSend;
    target[WATCH_FLAG] = false;
  };
}

interface IAssignedPaint {
  itemId: number;
  fieldName: string;
  code: string;
}

const assignedPaints = new Map<number, IAssignedPaint>();
let paintObserver: MutationObserver | undefined;
let paintScheduled = false;
let refreshArmed = false;
let listRefreshed = false;
let lastRefreshTry = 0;

export function clearAssignedPaints(): void {
  assignedPaints.clear();
  refreshArmed = false;
  listRefreshed = false;
  paintScheduled = false;
  lastRefreshTry = 0;
  if (paintObserver) {
    paintObserver.disconnect();
    paintObserver = undefined;
  }
}

export function paintNumbers(updates: IAssignedPaint[]): void {
  updates.forEach((update) => {
    assignedPaints.set(update.itemId, update);
  });
  watchAssignedRows();
  applyAssignedPaints();
}

export function repaintAssignedNumbers(): void {
  applyAssignedPaints();
}

export function clickListRefresh(): boolean {
  const named = document.querySelector(
    'button[data-automationid="refreshCommand"], button[data-id="Refresh"], button[name="Refresh"], button[aria-label="Refresh"], button[title="Refresh"]'
  );
  if (named instanceof HTMLButtonElement) {
    named.click();
    return true;
  }
  const buttons = document.querySelectorAll('button');
  for (let index = 0; index < buttons.length; index += 1) {
    const label = `${buttons[index].getAttribute('aria-label') || ''} ${buttons[index].getAttribute('title') || ''} ${buttons[index].getAttribute('name') || ''}`.toLowerCase();
    if (label.indexOf('refresh') >= 0) {
      buttons[index].click();
      return true;
    }
  }
  const icon = document.querySelector('[data-icon-name="Refresh"]');
  const button = icon ? icon.closest('button') : null;
  if (button instanceof HTMLButtonElement) {
    button.click();
    return true;
  }
  return false;
}

export function revealAssignedNumbers(itemIds: number[]): void {
  if (itemIds.length === 0) {
    return;
  }
  refreshArmed = true;
  listRefreshed = false;
  watchAssignedRows();
  applyAssignedPaints();
}

function watchAssignedRows(): void {
  if (paintObserver || assignedPaints.size === 0 || typeof MutationObserver === 'undefined' || !document.body) {
    return;
  }
  paintObserver = new MutationObserver(() => {
    if (paintScheduled) {
      return;
    }
    paintScheduled = true;
    window.requestAnimationFrame(() => {
      paintScheduled = false;
      applyAssignedPaints();
    });
  });
  paintObserver.observe(document.body, { childList: true, subtree: true, characterData: true });
}

function applyAssignedPaints(): void {
  assignedPaints.forEach((update) => {
    paintListCell(update);
    paintOpenForm(update);
  });
  maybeRefreshList();
}

function paintListCell(update: IAssignedPaint): void {
  const row = findRow(update.itemId);
  if (!row) {
    return;
  }
  const cell = findCell(row, update.fieldName);
  if (!cell) {
    return;
  }
  const input = cell.querySelector('input, textarea');
  if (input instanceof HTMLInputElement || input instanceof HTMLTextAreaElement) {
    writeControl(input, update.code);
    return;
  }
  const renderer = cell.querySelector('[data-automationid^="FieldRenderer"]') || cell.querySelector('span') || cell;
  if (visibleText(renderer)) {
    return;
  }
  renderer.textContent = update.code;
}

function paintOpenForm(update: IAssignedPaint): void {
  if (assignedPaints.size !== 1) {
    return;
  }
  const panel = document.querySelector('.ms-Panel-main');
  if (!panel) {
    return;
  }
  const control = findFormControl(panel, update.fieldName);
  if (!control) {
    return;
  }
  writeControl(control, update.code);
}

function writeControl(control: HTMLInputElement | HTMLTextAreaElement, code: string): void {
  if (control.value.trim()) {
    return;
  }
  control.value = code;
}

function findCell(row: Element, fieldName: string): Element | undefined {
  const name = fieldName.replace(/"/g, '');
  return row.querySelector(`[data-automation-key="${name}"]`)
    || row.querySelector(`[data-automationid="FieldRenderer-${name}"]`)
    || row.querySelector(`[data-field="${name}"]`)
    || undefined;
}

function findFormControl(root: ParentNode, fieldName: string): HTMLInputElement | HTMLTextAreaElement | undefined {
  const name = fieldName.replace(/"/g, '');
  const selectors = [
    `[data-automation-id="${name}"] input`,
    `[data-automation-id="${name}"] textarea`,
    `[data-automationid="${name}"] input`,
    `[data-automationid="${name}"] textarea`,
    `input[id*="${name}"]`,
    `textarea[id*="${name}"]`
  ];
  for (let index = 0; index < selectors.length; index += 1) {
    const node = root.querySelector(selectors[index]);
    if (node instanceof HTMLInputElement || node instanceof HTMLTextAreaElement) {
      return node;
    }
  }
  return undefined;
}

function visibleText(element: Element): string {
  return (element.textContent || '').replace(/\s+/g, ' ').trim();
}

function maybeRefreshList(): void {
  if (!refreshArmed || listRefreshed || gridEditIsOpen() || document.querySelector('.ms-Panel-main')) {
    return;
  }
  const now = Date.now();
  if (now - lastRefreshTry < 500) {
    return;
  }
  lastRefreshTry = now;
  listRefreshed = clickListRefresh();
}

export function gridEditIsOpen(): boolean {
  const buttons = document.querySelectorAll('button');
  for (let index = 0; index < buttons.length; index += 1) {
    const button = buttons[index];
    const label = `${button.getAttribute('aria-label') || ''} ${button.getAttribute('name') || ''}`.toLowerCase();
    if (label.indexOf('grid view') >= 0 && button.getAttribute('aria-pressed') === 'true') {
      return true;
    }
  }
  return false;
}

async function notifyIfCreated(response: Response, onCreated: () => Promise<void>): Promise<void> {
  try {
    const payload: unknown = await response.json();
    if (!createdItemId(payload)) {
      return;
    }
    await onCreated();
  } catch {
    // The item is already saved. The list check will number it.
  }
}

function findRow(itemId: number): Element | undefined {
  const marked = document.querySelector(`[role="row"][data-id="${itemId}"], [role="row"][data-item-key="${itemId}"]`);
  if (marked) {
    return marked;
  }
  const link = document.querySelector(`a[href*="ID=${itemId}"], a[href*="ID%3D${itemId}"], a[href*="ID%3d${itemId}"]`);
  if (!link) {
    return undefined;
  }
  return link.closest('[role="row"]') || undefined;
}

function requestUrl(input: RequestInfo | URL): string {
  if (typeof input === 'string') {
    return input;
  }
  if (input instanceof URL) {
    return input.toString();
  }
  return input.url;
}

function asRecord(value: unknown): { [key: string]: unknown } {
  return value && typeof value === 'object' ? value as { [key: string]: unknown } : {};
}

function arrayOf(value: unknown): unknown[] {
  return Array.isArray(value) ? value : [];
}

function positiveId(value: unknown): number | undefined {
  const id = typeof value === 'number' ? value : parseInt(String(value || ''), 10);
  return Number.isFinite(id) && id > 0 ? id : undefined;
}
