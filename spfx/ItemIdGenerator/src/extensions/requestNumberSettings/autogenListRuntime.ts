import * as React from 'react';
import * as ReactDom from 'react-dom';
import { SPHttpClient } from '@microsoft/sp-http';
import { buildListAbsoluteUrl, findConfigForList, isConfigList } from '../../webparts/itemIdGenerator/services/configContract';
import { connectConfiguredList, ensureTargetNumberColumn, loadConfigRows } from '../../webparts/itemIdGenerator/services/configListClient';
import { paintNumbers, revealAssignedNumbers, watchNewListItems } from '../../webparts/itemIdGenerator/services/listCreateWatch';
import { IAssignedNumber, numberBlankItems } from '../../webparts/itemIdGenerator/services/listNumbering';
import RequestNumberSettingsPanel from './RequestNumberSettingsPanel';

export const AUTOGEN_COMMAND_LABEL = 'Autogen Feature';
const BUTTON_ID = 'autogen-setting-command';
const STYLE_ID = 'autogen-setting-command-style';

export interface IAutogenListTarget {
  spHttpClient: SPHttpClient;
  siteAbsoluteUrl: string;
  webAbsoluteUrl: string;
  configListTitle: string;
  listId: string;
  listTitle: string;
  listServerRelativeUrl: string;
}

let commandSetTarget: IAutogenListTarget | undefined;
let customizerTarget: IAutogenListTarget | undefined;
let target: IAutogenListTarget | undefined;
let activeListId = '';
const numberGate = { busy: false, pending: false };
let timer: number | undefined;
let stopWatch: (() => void) | undefined;
let panelHost: HTMLDivElement | undefined;

export function syncAutogenFromCommandSet(next: IAutogenListTarget | undefined): void {
  commandSetTarget = usableTarget(next);
  applyActiveTarget();
}

export function syncAutogenFromCustomizer(next: IAutogenListTarget | undefined): void {
  customizerTarget = usableTarget(next);
  applyActiveTarget();
}

function usableTarget(next: IAutogenListTarget | undefined): IAutogenListTarget | undefined {
  if (!next || isConfigList(next.listTitle, next.listServerRelativeUrl, next.configListTitle)) {
    return undefined;
  }
  return next;
}

function applyActiveTarget(): void {
  const next = commandSetTarget || customizerTarget;
  if (!next) {
    stopRuntime();
    return;
  }
  target = next;
  if (activeListId !== next.listId) {
    stopWork();
    activeListId = next.listId;
    startWork(next);
  }
  ensureButton();
}

export function openAutogenSettingsPanel(): void {
  if (!target) {
    return;
  }
  if (!panelHost) {
    panelHost = document.body.appendChild(document.createElement('div'));
  }
  const current = target;
  const element: React.ReactElement = React.createElement(RequestNumberSettingsPanel, {
    siteAbsoluteUrl: current.siteAbsoluteUrl,
    configListTitle: current.configListTitle,
    listTitle: current.listTitle,
    listUrl: buildListAbsoluteUrl(current.webAbsoluteUrl, current.listServerRelativeUrl),
    listGuid: current.listId,
    spHttpClient: current.spHttpClient,
    onDismiss: closePanel
  });
  ReactDom.render(element, panelHost);
}

function startWork(next: IAutogenListTarget): void {
  stopWatch = watchNewListItems(next.listId, next.listServerRelativeUrl, () => numberList());
  showNumberColumn(next).catch(() => undefined);
  connectList(next).catch(() => undefined);
  numberList().catch(() => undefined);
  timer = window.setInterval(() => {
    ensureButton();
    numberList().catch(() => undefined);
  }, 1000);
}

function stopRuntime(): void {
  stopWork();
  target = undefined;
  activeListId = '';
  removeButton();
  closePanel();
}

function stopWork(): void {
  numberGate.pending = false;
  if (stopWatch) {
    stopWatch();
    stopWatch = undefined;
  }
  if (timer !== undefined) {
    window.clearInterval(timer);
    timer = undefined;
  }
}

async function numberList(): Promise<void> {
  if (!target) {
    return;
  }
  if (numberGate.busy) {
    numberGate.pending = true;
    return;
  }
  const current = target;
  numberGate.busy = true;
  let assigned: IAssignedNumber[] = [];
  try {
    assigned = await numberBlankItems(
      current.spHttpClient,
      current.siteAbsoluteUrl,
      current.configListTitle,
      current.webAbsoluteUrl,
      current.listId
    );
  } catch {
    assigned = [];
  }
  finishNumbering(current, assigned);
}

function finishNumbering(current: IAutogenListTarget, assigned: IAssignedNumber[]): void {
  numberGate.busy = false;
  if (assigned.length > 0) {
    showAssignedNumbers(assigned);
  }
  if (!numberGate.pending || !target || target.listId !== current.listId) {
    numberGate.pending = false;
    return;
  }
  numberGate.pending = false;
  numberList().catch(() => undefined);
}

async function showNumberColumn(next: IAutogenListTarget): Promise<void> {
  const rows = await loadConfigRows(next.spHttpClient, next.siteAbsoluteUrl, next.configListTitle);
  const config = findConfigForList(rows, next.listId);
  if (!config || !config.isActive) {
    return;
  }
  await ensureTargetNumberColumn(
    next.spHttpClient,
    buildListAbsoluteUrl(next.webAbsoluteUrl, next.listServerRelativeUrl),
    next.listId,
    config.numberColumnInternalName
  );
}

async function connectList(next: IAutogenListTarget): Promise<void> {
  const key = `requestNumberConnected:${next.listId}`;
  try {
    if (window.sessionStorage.getItem(key) === '1') {
      return;
    }
  } catch {
    // Session storage can be blocked. Still try to connect.
  }
  const connected = await connectConfiguredList(
    next.spHttpClient,
    next.siteAbsoluteUrl,
    next.configListTitle,
    next.listId
  );
  if (!connected) {
    return;
  }
  try {
    window.sessionStorage.setItem(key, '1');
  } catch {
    // The connection itself succeeded.
  }
}

function ensureButton(): void {
  // SharePoint rebuilds the list command bar and deletes buttons placed inside it.
  ensureStyles();
  if (document.getElementById(BUTTON_ID)) {
    return;
  }
  document.body.appendChild(createButton());
}

function createButton(): HTMLButtonElement {
  const button = document.createElement('button');
  button.id = BUTTON_ID;
  button.type = 'button';
  button.setAttribute('aria-label', AUTOGEN_COMMAND_LABEL);
  const label = document.createElement('span');
  label.textContent = AUTOGEN_COMMAND_LABEL;
  button.appendChild(label);
  button.addEventListener('click', (event: MouseEvent) => {
    event.preventDefault();
    event.stopPropagation();
    openAutogenSettingsPanel();
  });
  return button;
}

function removeButton(): void {
  const button = document.getElementById(BUTTON_ID);
  if (button && button.parentElement) {
    button.parentElement.removeChild(button);
  }
}

function ensureStyles(): void {
  if (document.getElementById(STYLE_ID)) {
    return;
  }
  const style = document.createElement('style');
  style.id = STYLE_ID;
  style.textContent = [
    `#${BUTTON_ID}{position:fixed;right:24px;bottom:24px;z-index:1000000;display:inline-flex;align-items:center;`,
    'height:40px;padding:0 16px;border:0;border-radius:20px;background:#038387;color:#fff;cursor:pointer;',
    'box-shadow:0 4px 16px rgba(0,0,0,.28);font:600 14px/40px "Segoe UI","Segoe UI Web (West European)",sans-serif;}',
    `#${BUTTON_ID}:hover{background:#026d70;}`
  ].join('');
  document.head.appendChild(style);
}

function closePanel(): void {
  if (panelHost) {
    ReactDom.unmountComponentAtNode(panelHost);
  }
}

function showAssignedNumbers(assigned: IAssignedNumber[]): void {
  paintNumbers(assigned);
  revealAssignedNumbers(assigned.map((item) => item.itemId));
}
