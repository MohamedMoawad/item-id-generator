import * as React from 'react';
import * as ReactDom from 'react-dom';
import { SPHttpClient } from '@microsoft/sp-http';
import { buildListAbsoluteUrl, findConfigForList, isConfigList, normalizeGuid } from '../../webparts/itemIdGenerator/services/configContract';
import { ensureTargetNumberColumn, loadConfigRows } from '../../webparts/itemIdGenerator/services/configListClient';
import { clearAssignedPaints, ensureListCreateWatch, IPreparedCreate, paintNumbers, refreshOpenList, repaintAssignedNumbers, revealAssignedNumbers, watchNewListItems } from '../../webparts/itemIdGenerator/services/listCreateWatch';
import { embedNumberInCreate, IAssignedNumber, numberBlankItems } from '../../webparts/itemIdGenerator/services/listNumbering';
import RequestNumberSettingsPanel from './RequestNumberSettingsPanel';

export const AUTOGEN_COMMAND_LABEL = 'Autogen Feature';

export interface IAutogenListTarget {
  spHttpClient: SPHttpClient;
  siteAbsoluteUrl: string;
  webAbsoluteUrl: string;
  configListTitle: string;
  listId: string;
  listTitle: string;
  listServerRelativeUrl: string;
}

interface IAutogenWindow extends Window {
  __autogenNumberTimer?: number;
}

let commandSetTarget: IAutogenListTarget | undefined;
let customizerTarget: IAutogenListTarget | undefined;
let target: IAutogenListTarget | undefined;
let activeListId = '';
let ownsTimer = false;
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
  const nextId = normalizeGuid(next.listId) || next.listId;
  if (activeListId !== nextId) {
    stopWork();
    activeListId = nextId;
    startWork(next);
  }
  removeInjectedButton();
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
  stopWatch = watchNewListItems(
    next.listId,
    next.listServerRelativeUrl,
    (body) => embedOnCreate(body),
    (itemId) => numberAfterCreate(itemId)
  );
  showNumberColumn(next).catch(() => undefined);
  numberList(true).catch(() => undefined);
  removeInjectedButton();
  const host = window as IAutogenWindow;
  if (host.__autogenNumberTimer === undefined) {
    ownsTimer = true;
    host.__autogenNumberTimer = window.setInterval(() => {
      removeInjectedButton();
      ensureListCreateWatch();
      repaintAssignedNumbers();
      numberList().catch(() => undefined);
    }, 1000);
  }
}

function stopRuntime(): void {
  stopWork();
  target = undefined;
  activeListId = '';
  removeInjectedButton();
  closePanel();
}

function stopWork(): void {
  clearAssignedPaints();
  if (stopWatch) {
    stopWatch();
    stopWatch = undefined;
  }
  const host = window as IAutogenWindow;
  if (ownsTimer && host.__autogenNumberTimer !== undefined) {
    window.clearInterval(host.__autogenNumberTimer);
    host.__autogenNumberTimer = undefined;
    ownsTimer = false;
  }
}

function embedOnCreate(body: string): Promise<IPreparedCreate | undefined> {
  if (!target) {
    return Promise.resolve(undefined);
  }
  const current = target;
  return embedNumberInCreate(
    current.spHttpClient,
    current.siteAbsoluteUrl,
    current.configListTitle,
    current.webAbsoluteUrl,
    current.listId,
    current.listServerRelativeUrl,
    body
  );
}

async function numberAfterCreate(itemId?: number): Promise<void> {
  await numberList(true);
  if (itemId) {
    revealAssignedNumbers([itemId]);
  }
  refreshOpenList();
}

async function numberList(force = false): Promise<boolean> {
  if (!target) {
    return false;
  }
  const current = target;
  let assigned: IAssignedNumber[] = [];
  try {
    assigned = await numberBlankItems(
      current.spHttpClient,
      current.siteAbsoluteUrl,
      current.configListTitle,
      current.webAbsoluteUrl,
      current.listId,
      force
    );
  } catch {
    assigned = [];
  }
  if (assigned.length > 0 && target && target.listId === current.listId) {
    showAssignedNumbers(assigned);
    return true;
  }
  return false;
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

function removeInjectedButton(): void {
  const button = document.getElementById('autogen-setting-command');
  if (button && button.parentElement) {
    button.parentElement.removeChild(button);
  }
  const style = document.getElementById('autogen-setting-command-style');
  if (style && style.parentElement) {
    style.parentElement.removeChild(style);
  }
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
