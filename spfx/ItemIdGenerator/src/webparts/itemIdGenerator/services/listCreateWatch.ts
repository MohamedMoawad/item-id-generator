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
let refreshTimer: number | undefined;
let refreshGeneration = 0;

export function clearAssignedPaints(): void {
  assignedPaints.clear();
  paintScheduled = false;
  refreshGeneration += 1;
  if (refreshTimer !== undefined) {
    window.clearTimeout(refreshTimer);
    refreshTimer = undefined;
  }
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

const REFRESH_SELECTORS = [
  'button[data-automationid="refreshCommand"]',
  'button[data-id="Refresh"]',
  'button[name="Refresh"]',
  'button[aria-label="Refresh"]',
  'button[title="Refresh"]'
];

export function clickListRefresh(): boolean {
  const button = findRefreshButton(document);
  if (!button) {
    return false;
  }
  button.click();
  return true;
}

export function revealAssignedNumbers(itemIds: number[]): void {
  if (itemIds.length === 0) {
    return;
  }
  watchAssignedRows();
  applyAssignedPaints();
  scheduleAjaxRefresh(itemIds);
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

function scheduleAjaxRefresh(itemIds: number[]): void {
  const generation = refreshGeneration;
  let overflowOpened = false;
  const started = Date.now();
  const tick = (): void => {
    refreshTimer = undefined;
    if (generation !== refreshGeneration) {
      return;
    }
    if (gridEditIsOpen() || document.querySelector('.ms-Panel-main')) {
      if (Date.now() - started < 20000) {
        refreshTimer = window.setTimeout(tick, 400);
      }
      return;
    }
    if (clickListRefresh()) {
      window.setTimeout(() => {
        if (generation === refreshGeneration) {
          applyAssignedPaints();
        }
      }, 700);
      return;
    }
    if (!overflowOpened && openCommandOverflow()) {
      overflowOpened = true;
      refreshTimer = window.setTimeout(tick, 350);
      return;
    }
    if (Date.now() - started < 2500) {
      refreshTimer = window.setTimeout(tick, 400);
      return;
    }
    reloadListOnce(itemIds);
  };
  if (refreshTimer !== undefined) {
    window.clearTimeout(refreshTimer);
  }
  refreshTimer = window.setTimeout(tick, 400);
}

function findRefreshButton(root: ParentNode): HTMLButtonElement | undefined {
  for (let index = 0; index < REFRESH_SELECTORS.length; index += 1) {
    const nodes = root.querySelectorAll(REFRESH_SELECTORS[index]);
    for (let nodeIndex = 0; nodeIndex < nodes.length; nodeIndex += 1) {
      const button = nodes[nodeIndex];
      if (button instanceof HTMLButtonElement && button.getClientRects().length > 0) {
        return button;
      }
    }
  }
  const icons = root.querySelectorAll('[data-icon-name="Refresh"]');
  for (let index = 0; index < icons.length; index += 1) {
    const button = icons[index].closest('button');
    if (button instanceof HTMLButtonElement && button.getClientRects().length > 0) {
      return button;
    }
  }
  return undefined;
}

function openCommandOverflow(): boolean {
  const bars = document.querySelectorAll('[data-automationid="ListViewCommandBar"], [role="menubar"]');
  for (let index = 0; index < bars.length; index += 1) {
    const overflow = bars[index].querySelector(
      'button[data-automationid="overflowButton"], button[aria-label="More commands"], button[aria-label="More"], button[title="More commands"], button[title="More"]'
    );
    if (overflow instanceof HTMLButtonElement && overflow.getClientRects().length > 0) {
      overflow.click();
      return true;
    }
  }
  return false;
}

function reloadListOnce(itemIds: number[]): void {
  const key = `autogenReload:${itemIds.slice().sort((left, right) => left - right).join(',')}`;
  try {
    if (window.sessionStorage.getItem(key) === '1') {
      return;
    }
    window.sessionStorage.setItem(key, '1');
  } catch {
    window.location.reload();
    return;
  }
  window.location.reload();
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
