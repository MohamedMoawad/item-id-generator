import { AadHttpClient, AadHttpClientFactory } from '@microsoft/sp-http';
import { buildCreateFlowPayload, IListFlowTarget } from './siteFlow';
import {
  connectionListUrls,
  connectionPortalUrl,
  FlowConsentError,
  FlowRunState,
  FlowStartError,
  ICreatedConnection,
  IFlowConnection,
  IFlowEnvironment,
  IRawListedFlow,
  isConsentFailure,
  readConnectionNamesFromFlows,
  readConnections,
  readCreatedConnection,
  readEnvironments,
  readFailedAction,
  readFlowState,
  readLatestRun,
  sharePointConnectionBody,
  sharePointConnectionCreateUrls,
  sharePointConnectionPutUrl
} from './flowModel';

const FLOW_RESOURCE = 'https://service.flow.microsoft.com/';
const API_VERSION = '2016-11-01';

export interface IFlowContext {
  environments: IFlowEnvironment[];
  environment: IFlowEnvironment;
  connections: IFlowConnection[];
}

export interface IRemoteFlow {
  id: string;
  state: FlowRunState;
}

const jsonHeaders: { [name: string]: string } = {
  Accept: 'application/json',
  'Content-Type': 'application/json'
};

class FlowRequestError extends Error {
  public readonly status: number;

  public constructor(status: number, message: string) {
    super(message);
    this.name = 'FlowRequestError';
    this.status = status;
    Object.setPrototypeOf(this, FlowRequestError.prototype);
  }
}

export async function listEnvironments(factory: AadHttpClientFactory): Promise<IFlowEnvironment[]> {
  const client = await openClient(factory);
  const payload = await sendJson<Parameters<typeof readEnvironments>[0]>(
    client,
    'GET',
    `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments?api-version=${API_VERSION}`
  );
  return readEnvironments(payload);
}

export async function listConnections(factory: AadHttpClientFactory, environment: IFlowEnvironment): Promise<IFlowConnection[]> {
  const client = await openClient(factory);
  return listConnectionsWithClient(client, environment);
}

async function connectionsFromExistingFlows(client: AadHttpClient, environment: IFlowEnvironment): Promise<IFlowConnection[]> {
  const hosts = [environment.endpoint.replace(/\/+$/, ''), 'https://api.flow.microsoft.com'];
  for (let hostIndex = 0; hostIndex < hosts.length; hostIndex += 1) {
    const host = hosts[hostIndex];
    if (hostIndex > 0 && host === hosts[0]) {
      continue;
    }
    const listUrl = `${host}/providers/Microsoft.ProcessSimple/environments/${encodeURIComponent(environment.id)}/flows?api-version=${API_VERSION}&$top=50`;
    try {
      const listed = await sendJson<{ value?: IRawListedFlow[] }>(client, 'GET', listUrl);
      const direct = readConnectionNamesFromFlows(listed);
      if (direct.length > 0) {
        return direct;
      }
      const ids = (listed.value || []).map((flow) => flow.name || '').filter((name) => name.length > 0).slice(0, 15);
      const details = await Promise.all(ids.map((flowId) => {
        const detailUrl = `${host}/providers/Microsoft.ProcessSimple/environments/${encodeURIComponent(environment.id)}/flows/${encodeURIComponent(flowId)}?api-version=${API_VERSION}`;
        return sendJson<Parameters<typeof readConnectionNamesFromFlows>[0]>(client, 'GET', detailUrl).catch(() => undefined);
      }));
      const collected: IFlowConnection[] = [];
      const seen: { [name: string]: boolean } = {};
      details.forEach((detail) => {
        if (!detail) {
          return;
        }
        readConnectionNamesFromFlows(detail).forEach((connection) => {
          if (!seen[connection.name]) {
            seen[connection.name] = true;
            collected.push(connection);
          }
        });
      });
      if (collected.length > 0) {
        return collected;
      }
    } catch (error) {
      if (error instanceof FlowConsentError) {
        throw error;
      }
    }
  }
  return [];
}

async function resolvePowerAppsHost(client: AadHttpClient, environment: IFlowEnvironment): Promise<IFlowEnvironment> {
  if (environment.powerAppsEndpoint !== 'https://api.powerapps.com') {
    return environment;
  }
  try {
    const detail = await sendJson<IRawEnvironmentRecord>(
      client,
      'GET',
      `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/${encodeURIComponent(environment.id)}?api-version=${API_VERSION}`
    );
    const refreshed = readEnvironments({ value: [detail] })[0];
    return refreshed || environment;
  } catch {
    return environment;
  }
}

interface IRawEnvironmentRecord {
  name?: string;
  properties?: {
    displayName?: string;
    isDefault?: boolean;
    runtimeEndpoints?: { [name: string]: string };
  };
}

export interface ISyncedFlow {
  flowId: string;
  flowEnvironment: string;
  flowConnection: string;
  flowState: string;
  note: string;
  consentUrl: string;
}

export async function syncListFlow(
  factory: AadHttpClientFactory,
  target: IListFlowTarget,
  current: ISyncedFlow,
  active: boolean
): Promise<ISyncedFlow> {
  const environments = await listEnvironments(factory);
  const environment = environments.filter((item) => item.id === current.flowEnvironment)[0] || environments[0];
  if (!active) {
    if (current.flowId && environment) {
      const stopAt = environments.filter((item) => item.id === current.flowEnvironment)[0] || environment;
      try {
        await setRemoteFlowState(factory, stopAt, current.flowId, 'Stopped');
      } catch (error) {
        if (error instanceof FlowConsentError) {
          throw error;
        }
      }
    }
    return {
      flowId: current.flowId,
      flowEnvironment: current.flowEnvironment,
      flowConnection: current.flowConnection,
      flowState: current.flowId ? 'Stopped' : '',
      note: '',
      consentUrl: ''
    };
  }
  if (!environment) {
    return { ...current, note: 'Saved the rule. Power Automate did not return an environment for this account.', consentUrl: '' };
  }
  const client = await openClient(factory);
  const ensured = await ensureSharePointConnection(client, environment, current.flowConnection);
  const connection = ensured.connection;
  if (!connection) {
    return {
      flowId: current.flowId,
      flowEnvironment: environment.id,
      flowConnection: current.flowConnection,
      flowState: current.flowState,
      note: 'The SharePoint connection could not be created in this environment. Open Power Automate, add a SharePoint connection for this account, then save again.',
      consentUrl: connectionPortalUrl(environment.id)
    };
  }
  const sameEnvironment = current.flowEnvironment === environment.id;
  if (current.flowId && sameEnvironment) {
    try {
      const remote = await redeployRemoteFlow(factory, environment, current.flowId, target, connection.name);
      return {
        flowId: remote.id,
        flowEnvironment: environment.id,
        flowConnection: connection.name,
        flowState: 'Started',
        note: ensured.consentUrl ? 'Sign in to the SharePoint connection, then save again so numbering can start.' : '',
        consentUrl: ensured.consentUrl
      };
    } catch (error) {
      if (error instanceof FlowStartError) {
        return flowSavedWithoutStart(environment, connection.name, error, ensured.consentUrl);
      }
      throw error;
    }
  }
  if (current.flowId) {
    const previous = environments.filter((item) => item.id === current.flowEnvironment)[0];
    if (previous) {
      try {
        await setRemoteFlowState(factory, previous, current.flowId, 'Stopped');
      } catch (error) {
        if (error instanceof FlowConsentError) {
          throw error;
        }
      }
    }
  }
  try {
    const created = await createRemoteFlow(factory, environment, target, connection.name);
    return {
      flowId: created.id,
      flowEnvironment: environment.id,
      flowConnection: connection.name,
      flowState: 'Started',
      note: ensured.consentUrl ? 'Sign in to the SharePoint connection, then save again so numbering can start.' : '',
      consentUrl: ensured.consentUrl
    };
  } catch (error) {
    if (error instanceof FlowStartError) {
      return flowSavedWithoutStart(environment, connection.name, error, ensured.consentUrl);
    }
    throw error;
  }
}

async function ensureSharePointConnection(
  client: AadHttpClient,
  environment: IFlowEnvironment,
  savedName: string
): Promise<{ connection?: IFlowConnection; consentUrl: string }> {
  const listed = await listConnectionsWithClient(client, environment);
  const known = listed.filter((item) => item.name === savedName)[0] || listed[0];
  if (known) {
    return { connection: known, consentUrl: '' };
  }
  if (savedName) {
    return { connection: { name: savedName, displayName: 'SharePoint', accountName: '' }, consentUrl: '' };
  }
  const created = await createSharePointConnection(client, environment);
  if (!created) {
    return { consentUrl: '' };
  }
  return {
    connection: { name: created.name, displayName: created.displayName, accountName: created.accountName },
    consentUrl: created.consentUrl
  };
}

async function listConnectionsWithClient(client: AadHttpClient, environment: IFlowEnvironment): Promise<IFlowConnection[]> {
  const resolved = await resolvePowerAppsHost(client, environment);
  const urls = connectionListUrls(resolved);
  let consentOnFlowHost = false;
  let fallback: IFlowConnection[] = [];
  for (let index = 0; index < urls.length; index += 1) {
    const url = urls[index];
    try {
      const payload = await sendJson<Parameters<typeof readConnections>[0]>(client, 'GET', url);
      const scoped = url.indexOf('shared_sharepointonline') >= 0;
      const rows = readConnections(payload, scoped, resolved.id);
      if (rows.length > 0) {
        return rows;
      }
      const loose = readConnections(payload, scoped, '');
      if (loose.length > fallback.length) {
        fallback = loose;
      }
    } catch (error) {
      if (error instanceof FlowConsentError && isFlowHost(url)) {
        consentOnFlowHost = true;
      }
    }
  }
  if (fallback.length > 0) {
    return fallback;
  }
  const fromFlows = await connectionsFromExistingFlows(client, resolved);
  if (fromFlows.length > 0) {
    return fromFlows;
  }
  if (consentOnFlowHost) {
    throw new FlowConsentError();
  }
  return [];
}

async function createSharePointConnection(client: AadHttpClient, environment: IFlowEnvironment): Promise<ICreatedConnection | undefined> {
  const body = sharePointConnectionBody(environment.id);
  const urls = sharePointConnectionCreateUrls(environment);
  let consentOnFlowHost = false;
  for (let index = 0; index < urls.length; index += 1) {
    const url = urls[index];
    try {
      const payload = await sendJson<Parameters<typeof readCreatedConnection>[0]>(client, 'POST', url, body);
      const created = readCreatedConnection(payload);
      if (created) {
        return created;
      }
    } catch (error) {
      if (error instanceof FlowConsentError && isFlowHost(url)) {
        consentOnFlowHost = true;
      }
    }
  }
  const name = newSharePointConnectionName();
  const putUrl = sharePointConnectionPutUrl(environment, name);
  try {
    const payload = await sendJson<Parameters<typeof readCreatedConnection>[0]>(client, 'PUT', putUrl, body);
    const created = readCreatedConnection(payload);
    if (created) {
      return created;
    }
  } catch (error) {
    if (error instanceof FlowConsentError && isFlowHost(putUrl)) {
      consentOnFlowHost = true;
    }
  }
  if (consentOnFlowHost) {
    throw new FlowConsentError();
  }
  return undefined;
}

function newSharePointConnectionName(): string {
  const alphabet = 'abcdefghijklmnopqrstuvwxyz0123456789';
  let suffix = '';
  for (let index = 0; index < 10; index += 1) {
    suffix += alphabet.charAt(Math.floor(Math.random() * alphabet.length));
  }
  return `shared-sharepointonl-${suffix}`;
}

function flowSavedWithoutStart(environment: IFlowEnvironment, connectionName: string, error: FlowStartError, consentUrl: string): ISyncedFlow {
  return {
    flowId: error.flowId,
    flowEnvironment: environment.id,
    flowConnection: connectionName,
    flowState: 'Stopped',
    note: consentUrl
      ? 'The numbering flow was created. Sign in to the SharePoint connection, then save again.'
      : error.message,
    consentUrl: consentUrl || connectionPortalUrl(environment.id)
  };
}

export async function removeListFlow(factory: AadHttpClientFactory, flowId: string, environmentId: string): Promise<void> {
  if (!flowId) {
    return;
  }
  const environments = await listEnvironments(factory);
  const environment = environments.filter((item) => item.id === environmentId)[0] || environments[0];
  if (!environment) {
    return;
  }
  try {
    await setRemoteFlowState(factory, environment, flowId, 'Stopped');
  } catch (error) {
    if (error instanceof FlowConsentError) {
      throw error;
    }
  }
  const client = await openClient(factory);
  try {
    await sendJson(client, 'DELETE', flowUrl(environment, flowId));
  } catch (error) {
    if (error instanceof FlowRequestError && error.status === 404) {
      return;
    }
    if (error instanceof FlowConsentError) {
      throw error;
    }
  }
}

export async function loadFlowContext(factory: AadHttpClientFactory): Promise<IFlowContext> {
  const environments = await listEnvironments(factory);
  const environment = environments[0];
  if (!environment) {
    throw new Error('Power Automate did not return an environment for this account.');
  }
  return {
    environments,
    environment,
    connections: await listConnections(factory, environment)
  };
}

export async function readRemoteFlow(
  factory: AadHttpClientFactory,
  environment: IFlowEnvironment,
  flowId: string
): Promise<IRemoteFlow> {
  const client = await openClient(factory);
  try {
    const payload = await sendJson<{ name?: string; properties?: { state?: string } }>(
      client,
      'GET',
      flowUrl(environment, flowId)
    );
    return { id: payload.name || flowId, state: readFlowState(payload) };
  } catch (error) {
    if (error instanceof FlowRequestError && error.status === 404) {
      return { id: '', state: 'Missing' };
    }
    throw error;
  }
}

export async function createRemoteFlow(
  factory: AadHttpClientFactory,
  environment: IFlowEnvironment,
  target: IListFlowTarget,
  connectionName: string
): Promise<IRemoteFlow> {
  const client = await openClient(factory);
  const created = await sendJson<{ name?: string; properties?: { state?: string } }>(
    client,
    'POST',
    `${environment.endpoint}/providers/Microsoft.ProcessSimple/environments/${encodeURIComponent(environment.id)}/flows?api-version=${API_VERSION}`,
    buildCreateFlowPayload(target, connectionName)
  );
  const id = created.name || '';
  if (!id) {
    throw new Error('Power Automate did not return the new numbering flow.');
  }
  await startFlow(client, environment, id, 'The numbering flow was created and did not start.');
  return { id, state: 'Started' };
}

export async function redeployRemoteFlow(
  factory: AadHttpClientFactory,
  environment: IFlowEnvironment,
  flowId: string,
  target: IListFlowTarget,
  connectionName: string
): Promise<IRemoteFlow> {
  const client = await openClient(factory);
  try {
    await sendJson(client, 'POST', `${flowPath(environment, flowId)}/stop?api-version=${API_VERSION}`);
  } catch (error) {
    if (error instanceof FlowConsentError) {
      throw error;
    }
  }
  await sendJson(
    client,
    'PATCH',
    flowUrl(environment, flowId),
    buildCreateFlowPayload(target, connectionName)
  );
  await startFlow(client, environment, flowId, 'The numbering flow was updated and did not start.');
  return { id: flowId, state: 'Started' };
}

export async function readFlowProblem(
  factory: AadHttpClientFactory,
  environmentId: string,
  flowId: string
): Promise<string> {
  if (!flowId || !environmentId) {
    return '';
  }
  const environments = await listEnvironments(factory);
  const environment = environments.filter((item) => item.id === environmentId)[0] || environments[0];
  if (!environment) {
    return '';
  }
  const client = await openClient(factory);
  const runs = await sendJson<Parameters<typeof readLatestRun>[0]>(
    client,
    'GET',
    `${flowPath(environment, flowId)}/runs?api-version=${API_VERSION}&$top=10`
  );
  const latest = readLatestRun(runs);
  if (!latest) {
    return 'The numbering flow has not run. Save the rule again, then add a new item and leave the number column empty. For an item that is already blank, open it and save it once.';
  }
  if (latest.status === 'Succeeded' || latest.status === 'Running' || latest.status === 'Waiting' || latest.status === 'Cancelled') {
    return '';
  }
  let detail = latest.error;
  try {
    const actions = await sendJson<Parameters<typeof readFailedAction>[0]>(
      client,
      'GET',
      `${flowPath(environment, flowId)}/runs/${encodeURIComponent(latest.name)}/actions?api-version=${API_VERSION}`
    );
    const failed = readFailedAction(actions);
    if (failed) {
      detail = failed;
    }
  } catch (error) {
    if (error instanceof FlowConsentError) {
      throw error;
    }
  }
  const reason = detail || 'Open that flow in Power Automate and read the failed run.';
  return `The numbering flow ran and did not write a number. ${reason} Save the rule again, then add a new item or open the blank item and save it.`;
}

export async function setRemoteFlowState(
  factory: AadHttpClientFactory,
  environment: IFlowEnvironment,
  flowId: string,
  state: 'Started' | 'Stopped'
): Promise<void> {
  const client = await openClient(factory);
  const action = state === 'Started' ? 'start' : 'stop';
  await sendJson(client, 'POST', `${flowPath(environment, flowId)}/${action}?api-version=${API_VERSION}`);
}

async function startFlow(client: AadHttpClient, environment: IFlowEnvironment, flowId: string, fallback: string): Promise<void> {
  try {
    await sendJson(client, 'POST', `${flowPath(environment, flowId)}/start?api-version=${API_VERSION}`);
  } catch (error) {
    if (error instanceof FlowConsentError) {
      throw error;
    }
    const message = error instanceof Error ? error.message : fallback;
    throw new FlowStartError(flowId, message);
  }
}

async function openClient(factory: AadHttpClientFactory): Promise<AadHttpClient> {
  try {
    return await factory.getClient(FLOW_RESOURCE);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    if (isConsentFailure(0, message)) {
      throw new FlowConsentError();
    }
    throw error instanceof Error ? error : new Error(message);
  }
}

async function sendJson<T>(client: AadHttpClient, method: string, url: string, body?: object): Promise<T> {
  const response = await client.fetch(url, AadHttpClient.configurations.v1, {
    method,
    headers: jsonHeaders,
    body: body === undefined ? undefined : JSON.stringify(body)
  });
  const text = await response.text();
  if (!response.ok) {
    if (isConsentFailure(response.status, text)) {
      throw new FlowConsentError();
    }
    throw new FlowRequestError(response.status, shorten(text, response.status));
  }
  if (!text.trim()) {
    return {} as T;
  }
  try {
    return JSON.parse(text) as T;
  } catch {
    throw new Error('Power Automate returned a response that was not JSON.');
  }
}

function isFlowHost(url: string): boolean {
  return url.indexOf('api.flow.microsoft.com') >= 0;
}

function flowPath(environment: IFlowEnvironment, flowId: string): string {
  return `${environment.endpoint}/providers/Microsoft.ProcessSimple/environments/${encodeURIComponent(environment.id)}/flows/${encodeURIComponent(flowId)}`;
}

function flowUrl(environment: IFlowEnvironment, flowId: string): string {
  return `${flowPath(environment, flowId)}?api-version=${API_VERSION}`;
}

function shorten(text: string, status: number): string {
  try {
    const payload = JSON.parse(text) as { error?: { message?: string }; message?: string };
    const message = (payload.error && payload.error.message) || payload.message;
    if (message) {
      return message.slice(0, 300);
    }
  } catch {
    // The body is plain text.
  }
  const detail = text.trim().slice(0, 300);
  return detail || `Power Automate returned HTTP ${status}.`;
}
