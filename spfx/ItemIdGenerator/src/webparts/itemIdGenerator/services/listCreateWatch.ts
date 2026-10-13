import { IFormulaColumn, isLookupLike, textFromFieldValue } from './formula';

const WATCH_FLAG = '__autoGenListCreateWatch';

export interface IPreparedCreate {
  body: string;
  rollback?: () => Promise<void>;
}

interface IWatchWindow extends Window {
  __autoGenListCreateWatch?: boolean;
  __autogenCreateListGuid?: string;
  __autogenCreateListUrl?: string;
  __autogenOnCreating?: (body: string) => Promise<IPreparedCreate | undefined>;
  __autogenOnCreated?: (itemId?: number) => Promise<void>;
}

interface IRecord {
  [key: string]: unknown;
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

export function readCreateField(body: string, fieldName: string): { found: boolean; value: string } | undefined {
  const fields = createFieldMap(body);
  if (!fields) {
    return undefined;
  }
  const wanted = fieldName.toLowerCase();
  if (!fields.has(wanted)) {
    return { found: false, value: '' };
  }
  return { found: true, value: fields.get(wanted) || '' };
}

export function embedCreateField(body: string, fieldName: string, value: string): string | undefined {
  let payload: unknown;
  try {
    payload = JSON.parse(body);
  } catch {
    return undefined;
  }
  if (!payload || typeof payload !== 'object' || Array.isArray(payload)) {
    return undefined;
  }
  const root = asRecord(payload);
  if (!stampCreateField(root, fieldName, value)) {
    return undefined;
  }
  return JSON.stringify(payload);
}

export function columnValuesFromCreateBody(
  body: string,
  names: string[],
  columns: IFormulaColumn[]
): { values: { [token: string]: string }; complete: boolean } | undefined {
  const fields = createFieldMap(body);
  if (!fields) {
    return undefined;
  }
  const values: { [token: string]: string } = {};
  let complete = true;
  names.forEach((name) => {
    const raw = fields.get(name.toLowerCase());
    if (raw === undefined) {
      complete = false;
      return;
    }
    const column = columns.filter((item) => item.internalName.toLowerCase() === name.toLowerCase())[0];
    const lookup = !!column && isLookupLike(column.type);
    const parsed = interpretCreateValue(raw, lookup || (!column && bareLookupId(raw)));
    if ((lookup || bareLookupId(raw)) && !parsed.titled) {
      complete = false;
    }
    values[name] = lookup ? (parsed.title || parsed.plain) : parsed.plain;
    values[`${name}.title`] = parsed.title || parsed.plain;
    values[`${name}.Title`] = values[`${name}.title`];
    values[`${name}.id`] = parsed.id;
    values[`${name}.Id`] = parsed.id;
  });
  return { values, complete };
}

export function newItemJson(body: string, listGuid: string, serverRelativeUrl: string): string | undefined {
  return locateNewItemJson(body, listGuid, serverRelativeUrl)?.json;
}

export function rewriteNewItemBody(
  body: string,
  listGuid: string,
  serverRelativeUrl: string,
  rewrite: (json: string) => string | undefined
): string | undefined {
  const span = locateNewItemJson(body, listGuid, serverRelativeUrl);
  if (!span) {
    return undefined;
  }
  const next = rewrite(span.json);
  if (!next) {
    return undefined;
  }
  try {
    JSON.parse(next);
  } catch {
    return undefined;
  }
  if (span.headerStart < 0) {
    return body.slice(0, span.start) + next + body.slice(span.end);
  }
  const headers = updateContentLength(body.slice(span.headerStart, span.headerEnd), byteLength(next));
  return body.slice(0, span.headerStart) + headers + body.slice(span.headerEnd, span.start) + next + body.slice(span.end);
}

export function createSaveFailed(payload: unknown): boolean {
  const rows = resultRows(payload);
  for (let index = 0; index < rows.length; index += 1) {
    const row = rows[index];
    if (row.HasException === true) {
      return true;
    }
    if (typeof row.ErrorMessage === 'string' && row.ErrorMessage.trim()) {
      return true;
    }
  }
  return false;
}

export function batchCreateOutcome(body: string): 'saved' | 'failed' | 'unknown' {
  const trimmed = body.trim();
  if (!trimmed) {
    return 'unknown';
  }
  if (trimmed.startsWith('{') || trimmed.startsWith('[')) {
    try {
      const payload: unknown = JSON.parse(trimmed);
      if (createSaveFailed(payload)) {
        return 'failed';
      }
      return createdItemId(payload) ? 'saved' : 'unknown';
    } catch {
      return 'unknown';
    }
  }
  let saved = false;
  let failed = false;
  let cursor = 0;
  while (cursor < body.length) {
    const start = body.indexOf('{', cursor);
    if (start < 0) {
      break;
    }
    const extracted = extractJsonObject(body, start);
    if (!extracted) {
      break;
    }
    try {
      const payload: unknown = JSON.parse(extracted.json);
      if (createSaveFailed(payload)) {
        failed = true;
      }
      if (createdItemId(payload)) {
        saved = true;
      }
    } catch {
      // Keep scanning. Batch parts can contain braces that are not the item payload.
    }
    cursor = extracted.end;
  }
  if (saved) {
    return 'saved';
  }
  if (failed) {
    return 'failed';
  }
  const statuses = body.match(/HTTP\/\d+(?:\.\d+)?[ \t]+(\d{3})/g) || [];
  for (let index = 0; index < statuses.length; index += 1) {
    const code = parseInt(statuses[index].replace(/\D/g, ''), 10);
    if (code >= 400) {
      return 'failed';
    }
  }
  return 'unknown';
}

export function watchNewListItems(
  listGuid: string,
  serverRelativeUrl: string,
  onCreating: (body: string) => Promise<IPreparedCreate | undefined>,
  onCreated: (itemId?: number) => Promise<void>
): () => void {
  const target = window as IWatchWindow;
  target.__autogenCreateListGuid = listGuid;
  target.__autogenCreateListUrl = serverRelativeUrl;
  target.__autogenOnCreating = onCreating;
  target.__autogenOnCreated = onCreated;
  const releaseCallbacks = (): void => {
    if (target.__autogenOnCreating === onCreating) {
      target.__autogenOnCreating = undefined;
    }
    if (target.__autogenOnCreated === onCreated) {
      target.__autogenOnCreated = undefined;
    }
  };
  if (target[WATCH_FLAG]) {
    return releaseCallbacks;
  }
  target[WATCH_FLAG] = true;
  const originalFetch = window.fetch.bind(window);
  const originalOpen = XMLHttpRequest.prototype.open;
  const originalSend = XMLHttpRequest.prototype.send;

  window.fetch = (input: RequestInfo | URL, init?: RequestInit): Promise<Response> => {
    const url = requestUrl(input);
    const method = requestMethod(input, init);
    if (!mayCarryNewItem(url, method)) {
      return originalFetch(input, init);
    }
    return sendPreparedFetch(originalFetch, input, init);
  };

  XMLHttpRequest.prototype.open = function (this: XMLHttpRequest, method: string, url: string | URL): void {
    (this as XMLHttpRequest & { __autoGenMethod?: string; __autoGenUrl?: string }).__autoGenMethod = method;
    (this as XMLHttpRequest & { __autoGenUrl?: string }).__autoGenUrl = String(url);
    return originalOpen.apply(this, arguments as unknown as Parameters<typeof originalOpen>);
  };

  XMLHttpRequest.prototype.send = function (this: XMLHttpRequest): void {
    const marked = this as XMLHttpRequest & { __autoGenMethod?: string; __autoGenUrl?: string };
    const sent = arguments[0] as unknown;
    const textBody = typeof sent === 'string' ? sent : undefined;
    const url = marked.__autoGenUrl || '';
    const method = marked.__autoGenMethod || 'GET';
    if (textBody === undefined || !shouldInspectCreate(url, method, textBody)) {
      return originalSend.apply(this, arguments as unknown as Parameters<typeof originalSend>);
    }
    sendPreparedXhr(this, textBody, originalSend).catch(() => undefined);
  };

  return () => {
    releaseCallbacks();
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

function mayCarryNewItem(url: string, method: string): boolean {
  const host = window as IWatchWindow;
  const listGuid = host.__autogenCreateListGuid || '';
  const listUrl = host.__autogenCreateListUrl || '';
  if (isNewItemRequest(url, method, listGuid, listUrl)) {
    return true;
  }
  if ((method || '').toUpperCase() !== 'POST') {
    return false;
  }
  const lower = safeDecode(url).toLowerCase();
  return lower.indexOf('$batch') >= 0 || lower.indexOf('%24batch') >= 0;
}

function shouldInspectCreate(url: string, method: string, body: string): boolean {
  const host = window as IWatchWindow;
  const listGuid = host.__autogenCreateListGuid || '';
  const listUrl = host.__autogenCreateListUrl || '';
  if ((method || '').toUpperCase() !== 'POST') {
    return false;
  }
  if (isNewItemRequest(url, method, listGuid, listUrl)) {
    return true;
  }
  if (!mayCarryNewItem(url, method)) {
    return false;
  }
  const decoded = safeDecode(body).toLowerCase();
  const onList = (listGuid !== '' && decoded.indexOf(listGuid.toLowerCase()) >= 0) ||
    (listUrl.length > 1 && decoded.indexOf(listUrl.toLowerCase()) >= 0);
  return onList && (decoded.indexOf('addvalidateupdateitemusingpath') >= 0 || /\/items(\?|["'\s]|$)/i.test(decoded));
}

async function sendPreparedFetch(
  originalFetch: typeof window.fetch,
  input: RequestInfo | URL,
  init?: RequestInit
): Promise<Response> {
  const url = requestUrl(input);
  const method = requestMethod(input, init);
  const text = await readBodyText(input, init);
  if (text === undefined || !shouldInspectCreate(url, method, text)) {
    return originalFetch(input, init);
  }
  let nextInput: RequestInfo | URL = input;
  let nextInit = init;
  let rollback: (() => Promise<void>) | undefined;
  const prepared = await prepareCreateBody(text);
  if (prepared && prepared.body !== text) {
    const rebuilt = fetchWithBody(input, init, prepared.body);
    nextInput = rebuilt.input;
    nextInit = rebuilt.init;
    rollback = prepared.rollback;
  }
  try {
    const response = await originalFetch(nextInput, nextInit);
    const responseText = await response.clone().text().catch(() => '');
    await settleCreate(response.ok, responseText, rollback);
    return response;
  } catch (error) {
    await runRollback(rollback);
    throw error;
  }
}

async function sendPreparedXhr(
  xhr: XMLHttpRequest,
  body: string,
  originalSend: typeof XMLHttpRequest.prototype.send
): Promise<void> {
  let nextBody = body;
  let rollback: (() => Promise<void>) | undefined;
  const prepared = await prepareCreateBody(body);
  if (prepared && prepared.body !== body) {
    nextBody = prepared.body;
    rollback = prepared.rollback;
  }
  xhr.addEventListener('load', () => {
    const responseText = xhrResponseText(xhr);
    settleCreate(xhr.status >= 200 && xhr.status < 300, responseText, rollback).catch(() => undefined);
  });
  xhr.addEventListener('error', () => {
    runRollback(rollback).catch(() => undefined);
  });
  xhr.addEventListener('abort', () => {
    runRollback(rollback).catch(() => undefined);
  });
  originalSend.call(xhr, nextBody);
}

async function prepareCreateBody(body: string): Promise<IPreparedCreate | undefined> {
  const creating = (window as IWatchWindow).__autogenOnCreating;
  if (!creating) {
    return undefined;
  }
  try {
    return await creating(body);
  } catch {
    return undefined;
  }
}

async function settleCreate(ok: boolean, responseText: string, rollback?: () => Promise<void>): Promise<void> {
  if (!ok) {
    await runRollback(rollback);
    return;
  }
  const outcome = batchCreateOutcome(responseText);
  if (outcome === 'failed') {
    await runRollback(rollback);
    return;
  }
  if (outcome !== 'saved') {
    return;
  }
  const created = (window as IWatchWindow).__autogenOnCreated;
  if (!created) {
    return;
  }
  try {
    await created(savedItemId(responseText));
  } catch {
    // The item is already saved. The list check will number a blank column.
  }
}

function savedItemId(body: string): number | undefined {
  const trimmed = body.trim();
  if (!trimmed) {
    return undefined;
  }
  if (trimmed.startsWith('{') || trimmed.startsWith('[')) {
    try {
      return createdItemId(JSON.parse(trimmed));
    } catch {
      return undefined;
    }
  }
  let cursor = 0;
  while (cursor < body.length) {
    const start = body.indexOf('{', cursor);
    if (start < 0) {
      return undefined;
    }
    const extracted = extractJsonObject(body, start);
    if (!extracted) {
      return undefined;
    }
    try {
      const id = createdItemId(JSON.parse(extracted.json));
      if (id) {
        return id;
      }
    } catch {
      // Keep scanning the batch response.
    }
    cursor = extracted.end;
  }
  return undefined;
}

async function runRollback(rollback?: () => Promise<void>): Promise<void> {
  if (!rollback) {
    return;
  }
  try {
    await rollback();
  } catch {
    // The sequence can gap when the counter cannot be put back.
  }
}

function requestMethod(input: RequestInfo | URL, init?: RequestInit): string {
  if (init && init.method) {
    return init.method;
  }
  if (typeof Request !== 'undefined' && input instanceof Request) {
    return input.method;
  }
  return 'GET';
}

async function readBodyText(input: RequestInfo | URL, init?: RequestInit): Promise<string | undefined> {
  if (init && typeof init.body === 'string') {
    return init.body;
  }
  if (typeof Request !== 'undefined' && input instanceof Request) {
    try {
      return await input.clone().text();
    } catch {
      return undefined;
    }
  }
  return undefined;
}

function fetchWithBody(input: RequestInfo | URL, init: RequestInit | undefined, body: string): { input: string; init: RequestInit } {
  const headers = new Headers(typeof Request !== 'undefined' && input instanceof Request ? input.headers : undefined);
  if (init && init.headers) {
    new Headers(init.headers).forEach((value, key) => headers.set(key, value));
  }
  const request = typeof Request !== 'undefined' && input instanceof Request ? input : undefined;
  return {
    input: requestUrl(input),
    init: {
      method: requestMethod(input, init),
      headers,
      body,
      credentials: init && init.credentials ? init.credentials : request && request.credentials,
      cache: init && init.cache,
      redirect: init && init.redirect,
      referrer: init && init.referrer,
      integrity: init && init.integrity,
      mode: init && init.mode ? init.mode : request && request.mode,
      signal: init && init.signal ? init.signal : request && request.signal
    }
  };
}

function xhrResponseText(xhr: XMLHttpRequest): string {
  if (xhr.responseType === '' || xhr.responseType === 'text') {
    return xhr.responseText || '';
  }
  if (xhr.responseType === 'json' && xhr.response) {
    try {
      return JSON.stringify(xhr.response);
    } catch {
      return '';
    }
  }
  return '';
}

interface IJsonSpan {
  json: string;
  start: number;
  end: number;
  headerStart: number;
  headerEnd: number;
}

function locateNewItemJson(body: string, listGuid: string, serverRelativeUrl: string): IJsonSpan | undefined {
  const trimmed = body.trim();
  if (trimmed.startsWith('{')) {
    try {
      const payload: unknown = JSON.parse(trimmed);
      if (payload && typeof payload === 'object' && !Array.isArray(payload)) {
        const start = body.indexOf('{');
        const extracted = extractJsonObject(body, start);
        if (extracted) {
          return { json: extracted.json, start, end: extracted.end, headerStart: -1, headerEnd: -1 };
        }
      }
    } catch {
      // A batch body can still contain a create after this.
    }
  }
  return locateBatchJson(body, listGuid, serverRelativeUrl);
}

function locateBatchJson(body: string, listGuid: string, serverRelativeUrl: string): IJsonSpan | undefined {
  const pattern = /^POST[ \t]+(\S+)[ \t]+HTTP\/\d+(?:\.\d+)?[ \t]*\r?$/gim;
  let match = pattern.exec(body);
  while (match) {
    const request = match[1];
    const afterLine = match.index + match[0].length;
    match = pattern.exec(body);
    if (!isNewItemRequest(request, 'POST', listGuid, serverRelativeUrl)) {
      continue;
    }
    const headerBreak = findHeaderBreak(body, afterLine);
    if (!headerBreak) {
      continue;
    }
    const start = skipSpace(body, headerBreak.bodyStart);
    const extracted = extractJsonObject(body, start);
    if (!extracted) {
      continue;
    }
    return {
      json: extracted.json,
      start: extracted.start,
      end: extracted.end,
      headerStart: afterLine,
      headerEnd: headerBreak.headerEnd
    };
  }
  return undefined;
}

function findHeaderBreak(text: string, from: number): { headerEnd: number; bodyStart: number } | undefined {
  const crlf = text.indexOf('\r\n\r\n', from);
  const lf = text.indexOf('\n\n', from);
  if (crlf >= 0 && (lf < 0 || crlf <= lf)) {
    return { headerEnd: crlf, bodyStart: crlf + 4 };
  }
  if (lf >= 0) {
    return { headerEnd: lf, bodyStart: lf + 2 };
  }
  return undefined;
}

function extractJsonObject(text: string, start: number): { json: string; start: number; end: number } | undefined {
  if (text.charAt(start) !== '{') {
    return undefined;
  }
  let depth = 0;
  let inString = false;
  let escaped = false;
  for (let index = start; index < text.length; index += 1) {
    const char = text.charAt(index);
    if (inString) {
      if (escaped) {
        escaped = false;
        continue;
      }
      if (char === '\\') {
        escaped = true;
        continue;
      }
      if (char === '"') {
        inString = false;
      }
      continue;
    }
    if (char === '"') {
      inString = true;
      continue;
    }
    if (char === '{') {
      depth += 1;
    }
    if (char === '}') {
      depth -= 1;
      if (depth === 0) {
        return { json: text.slice(start, index + 1), start, end: index + 1 };
      }
    }
  }
  return undefined;
}

function updateContentLength(headers: string, length: number): string {
  if (!/content-length[ \t]*:/i.test(headers)) {
    return headers;
  }
  return headers.replace(/content-length[ \t]*:[ \t]*\d+/i, `Content-Length: ${length}`);
}

function byteLength(value: string): number {
  if (typeof TextEncoder !== 'undefined') {
    return new TextEncoder().encode(value).length;
  }
  return value.length;
}

function skipSpace(text: string, index: number): number {
  let cursor = index;
  while (cursor < text.length && ' \t\r\n'.indexOf(text.charAt(cursor)) >= 0) {
    cursor += 1;
  }
  return cursor;
}

function createFieldMap(body: string): Map<string, string> | undefined {
  let payload: unknown;
  try {
    payload = JSON.parse(body);
  } catch {
    return undefined;
  }
  const root = asRecord(payload);
  if (!Object.keys(root).length) {
    return undefined;
  }
  const rows = formRowList(root) || formRowList(asRecord(root.d));
  if (rows) {
    const map = new Map<string, string>();
    rows.forEach((row) => {
      const record = asRecord(row);
      const name = stringProp(record, 'FieldName');
      if (!name) {
        return;
      }
      map.set(name.toLowerCase(), fieldValueText(record));
    });
    return map;
  }
  if (containsFormValues(root)) {
    return undefined;
  }
  const item = itemRecord(root);
  if (!item) {
    return undefined;
  }
  const map = new Map<string, string>();
  Object.keys(item).forEach((key) => {
    if (key === '__metadata' || key.indexOf('@') >= 0 || key.toLowerCase().indexOf('odata') >= 0) {
      return;
    }
    map.set(key.toLowerCase(), plainFieldText(item[key]));
  });
  return map;
}

function stampCreateField(root: IRecord, fieldName: string, value: string): boolean {
  const rows = formRowList(root) || formRowList(asRecord(root.d));
  if (rows) {
    return upsertFormRow(rows, fieldName, value);
  }
  if (containsFormValues(root)) {
    return false;
  }
  const item = itemRecord(root);
  if (!item) {
    return false;
  }
  item[fieldName] = value;
  return true;
}

function formRowList(record: IRecord): unknown[] | undefined {
  const direct = readFormRows(record);
  if (direct) {
    return direct;
  }
  const nested = asRecord(record.parameters) || asRecord(record.Parameters);
  return nested ? readFormRows(nested) : undefined;
}

function readFormRows(record: IRecord): unknown[] | undefined {
  const direct = record.formValues !== undefined ? record.formValues : record.FormValues;
  if (Array.isArray(direct)) {
    return direct;
  }
  const wrapped = asRecord(direct);
  if (Array.isArray(wrapped.results)) {
    return wrapped.results;
  }
  return undefined;
}

function containsFormValues(record: IRecord): boolean {
  return hasOwnFormValues(record) ||
    hasOwnFormValues(asRecord(record.parameters)) ||
    hasOwnFormValues(asRecord(record.Parameters)) ||
    hasOwnFormValues(asRecord(record.d));
}

function hasOwnFormValues(record: IRecord): boolean {
  return Object.prototype.hasOwnProperty.call(record, 'formValues') || Object.prototype.hasOwnProperty.call(record, 'FormValues');
}

function upsertFormRow(rows: unknown[], fieldName: string, value: string): boolean {
  const wanted = fieldName.toLowerCase();
  for (let index = 0; index < rows.length; index += 1) {
    const row = asRecord(rows[index]);
    if (!stringProp(row, 'FieldName') || stringProp(row, 'FieldName').toLowerCase() !== wanted) {
      continue;
    }
    if (Object.prototype.hasOwnProperty.call(row, 'fieldValue') && !Object.prototype.hasOwnProperty.call(row, 'FieldValue')) {
      row.fieldValue = value;
    } else {
      row.FieldValue = value;
    }
    return true;
  }
  rows.push({ FieldName: fieldName, FieldValue: value });
  return true;
}

function itemRecord(root: IRecord): IRecord | undefined {
  if (containsFormValues(root)) {
    return undefined;
  }
  const wrapped = asRecord(root.d);
  if (Object.keys(wrapped).length && !Array.isArray(wrapped.results)) {
    return wrapped;
  }
  if (Array.isArray(root.value) || Array.isArray(root.results)) {
    return undefined;
  }
  return root;
}

function fieldValueText(record: IRecord): string {
  const value = record.FieldValue !== undefined ? record.FieldValue : record.fieldValue;
  return plainFieldText(value);
}

function plainFieldText(value: unknown): string {
  if (value === null || value === undefined) {
    return '';
  }
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    return String(value);
  }
  return JSON.stringify(value);
}

function stringProp(record: IRecord, key: string): string {
  const camel = key.charAt(0).toLowerCase() + key.slice(1);
  const value = record[key] !== undefined ? record[key] : record[camel];
  return typeof value === 'string' ? value : '';
}

interface IInterpretedValue {
  plain: string;
  title: string;
  id: string;
  titled: boolean;
}

function interpretCreateValue(raw: string, lookup: boolean): IInterpretedValue {
  const trimmed = raw.trim();
  if (trimmed.startsWith('{') || trimmed.startsWith('[')) {
    try {
      const parsed: unknown = JSON.parse(trimmed);
      const title = textFromFieldValue(parsed, 'title') || textFromFieldValue(parsed);
      const plain = textFromFieldValue(parsed) || title;
      return {
        plain,
        title,
        id: textFromFieldValue(parsed, 'id'),
        titled: title.length > 0 && !bareLookupId(title)
      };
    } catch {
      // The field is ordinary text that happens to start with a brace.
    }
  }
  const classic = splitClassicLookup(trimmed);
  if (classic) {
    return {
      plain: classic.titles || classic.ids,
      title: classic.titles,
      id: classic.ids,
      titled: classic.titles.length > 0
    };
  }
  return {
    plain: trimmed,
    title: trimmed,
    id: '',
    titled: !lookup || !bareLookupId(trimmed)
  };
}

function splitClassicLookup(value: string): { ids: string; titles: string } | undefined {
  if (value.indexOf(';#') < 0) {
    return undefined;
  }
  const parts = value.split(';#');
  if (parts.length < 2 || parts.length % 2 !== 0) {
    return undefined;
  }
  const ids: string[] = [];
  const titles: string[] = [];
  for (let index = 0; index < parts.length; index += 2) {
    if (!/^\d+$/.test(parts[index])) {
      return undefined;
    }
    ids.push(parts[index]);
    if (parts[index + 1].trim()) {
      titles.push(parts[index + 1]);
    }
  }
  return { ids: ids.join(', '), titles: titles.join(', ') };
}

function bareLookupId(value: string): boolean {
  return /^\d+(?:;#\d+)*$/.test(value.trim());
}

function safeDecode(value: string): string {
  try {
    return decodeURIComponent(value);
  } catch {
    return value.replace(/%2f/ig, '/');
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

function asRecord(value: unknown): IRecord {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    return {};
  }
  return value as IRecord;
}

function arrayOf(value: unknown): unknown[] {
  return Array.isArray(value) ? value : [];
}

function resultRows(payload: unknown): IRecord[] {
  const record = asRecord(payload);
  const nested = asRecord(asRecord(record.d).AddValidateUpdateItemUsingPath);
  const sources = [record.value, nested.results, record.results, nested.value];
  for (let index = 0; index < sources.length; index += 1) {
    const rows = arrayOf(sources[index]);
    if (!rows.length) {
      continue;
    }
    const parsed: IRecord[] = [];
    rows.forEach((row) => {
      const recordRow = asRecord(row);
      if (Object.keys(recordRow).length) {
        parsed.push(recordRow);
      }
    });
    return parsed;
  }
  return [];
}

function positiveId(value: unknown): number | undefined {
  const id = typeof value === 'number' ? value : parseInt(String(value || ''), 10);
  return Number.isFinite(id) && id > 0 ? id : undefined;
}
