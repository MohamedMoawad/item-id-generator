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
  const clients = await openConnectionClients(factory);
  for (let index = 0; index < clients.ordered.length; index += 1) {
    const rows = await listConnectionsWithClient(clients.ordered[index], environment);
    if (rows.length > 0) {
      return rows;
    }
  }
  return [];
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
  const ensured = await ensureSharePointConnection(factory, environment, current.flowConnection);
  const names = ensured.connection ? [ensured.connection.name] : ['', newSharePointConnectionName()];
  let lastError = ensured.error;
  for (let index = 0; index < names.length; index += 1) {
    const connectionName = names[index];
    const sameEnvironment = current.flowEnvironment === environment.id;
    if (index === 0 && current.flowId && !sameEnvironment) {
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
      if (current.flowId && sameEnvironment) {
        const remote = await redeployRemoteFlow(factory, environment, current.flowId, target, connectionName);
        return startedFlow(environment, remote.id, connectionName, ensured.consentUrl);
      }
      const created = await createRemoteFlow(factory, environment, target, connectionName);
      return startedFlow(environment, created.id, connectionName, ensured.consentUrl);
    } catch (error) {
      if (error instanceof FlowStartError) {
        return flowSavedWithoutStart(environment, connectionName, error, ensured.consentUrl || connectionPortalUrl(environment.id));
      }
      if (error instanceof FlowConsentError) {
        throw error;
      }
      lastError = error instanceof Error ? error.message : lastError;
    }
  }
  return {
    flowId: current.flowId,
    flowEnvironment: environment.id,
    flowConnection: current.flowConnection,
    flowState: current.flowState,
    note: lastError || 'The numbering flow was not created.',
    consentUrl: connectionPortalUrl(environment.id)
  };
}

function startedFlow(environment: IFlowEnvironment, flowId: string, connectionName: string, consentUrl: string): ISyncedFlow {
  return {
    flowId,
    flowEnvironment: environment.id,
    flowConnection: connectionName,
    flowState: 'Started',
    note: consentUrl ? 'Sign in to the SharePoint connection, then save again so numbering can start.' : '',
    consentUrl
  };
}

async function ensureSharePointConnection(
  factory: AadHttpClientFactory,
  environment: IFlowEnvironment,
  savedName: string
): Promise<{ connection?: IFlowConnection; consentUrl: string; error: string }> {
  const clients = await openConnectionClients(factory);
  for (let index = 0; index < clients.ordered.length; index += 1) {
    const listed = await listConnectionsWithClient(clients.ordered[index], environment);
    const known = listed.filter((item) => item.name === savedName)[0] || listed[0];
    if (known) {
      return { connection: known, consentUrl: '', error: '' };
    }
  }
  let lastError = '';
  for (let index = 0; index < clients.ordered.length; index += 1) {
    const created = await createSharePointConnection(clients.ordered[index], environment);
    if (created.connection) {
      return {
        connection: {
          name: created.connection.name,
          displayName: created.connection.displayName,
          accountName: created.connection.accountName
        },
        consentUrl: created.connection.consentUrl,
        error: ''
      };
    }
    lastError = created.error || lastError;
  }
  if (savedName) {
    return { connection: { name: savedName, displayName: 'SharePoint', accountName: '' }, consentUrl: '', error: '' };
  }
  return { consentUrl: '', error: lastError };
}

async function listConnectionsWithClient(client: AadHttpClient, environment: IFlowEnvironment): Promise<IFlowConnection[]> {
  const resolved = await resolvePowerAppsHost(client, environment);
  const urls = connectionListUrls(resolved);
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
      if (!(error instanceof FlowConsentError) && !(error instanceof FlowRequestError)) {
        throw error;
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
  return [];
}

async function createSharePointConnection(
  client: AadHttpClient,
  environment: IFlowEnvironment
): Promise<{ connection?: ICreatedConnection; consent: boolean; error: string }> {
  const body = sharePointConnectionBody(environment.id);
  const urls = sharePointConnectionCreateUrls(environment);
  let consent = false;
  let errorText = '';
  for (let index = 0; index < urls.length; index += 1) {
    const url = urls[index];
    try {
      const payload = await sendJson<Parameters<typeof readCreatedConnection>[0]>(client, 'POST', url, body);
      const created = readCreatedConnection(payload);
      if (created) {
        return { connection: created, consent: false, error: '' };
      }
    } catch (error) {
      const failure = connectionFailure(error);
      consent = consent || failure.consent;
      errorText = failure.error || errorText;
    }
  }
  const name = newSharePointConnectionName();
  try {
    const payload = await sendJson<Parameters<typeof readCreatedConnection>[0]>(
      client,
      'PUT',
      sharePointConnectionPutUrl(environment, name),
      body
    );
    const created = readCreatedConnection(payload);
    if (created) {
      return { connection: created, consent: false, error: '' };
    }
  } catch (error) {
    const failure = connectionFailure(error);
    consent = consent || failure.consent;
    errorText = failure.error || errorText;
  }
  return { consent, error: errorText };
}

function connectionFailure(error: unknown): { consent: boolean; error: string } {
  if (error instanceof FlowConsentError) {
    return { consent: true, error: '' };
  }
  if (error instanceof FlowRequestError) {
    return { consent: false, error: error.message };
  }
  return { consent: false, error: error instanceof Error ? error.message : '' };
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
    `${flowApiRoot(environment)}/providers/Microsoft.ProcessSimple/environments/${encodeURIComponent(environment.id)}/flows?api-version=${API_VERSION}`,
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

async function openConnectionClients(factory: AadHttpClientFactory): Promise<{ powerApps?: AadHttpClient; ordered: AadHttpClient[] }> {
  const flow = await openClient(factory);
  return { ordered: [flow] };
}

async function openClient(factory: AadHttpClientFactory): Promise<AadHttpClient> {
  return factory.getClient(FLOW_RESOURCE);
}

async function sendJson<T>(client: AadHttpClient, method: string, url: string, body?: object): Promise<T> {
  const response = await client.fetch(url, AadHttpClient.configurations.v1, {
    method,
    headers: jsonHeaders,
    body: body === undefined ? undefined : JSON.stringify(body)
  });
  const text = await response.text();
  if (!response.ok) {
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

function flowApiRoot(environment: IFlowEnvironment): string {
  const endpoint = environment.endpoint.replace(/\/+$/, '');
  const match = /^https:\/\/([^/:?#]+)/i.exec(endpoint);
  const host = match ? match[1].toLowerCase() : '';
  if (host === 'api.flow.microsoft.com' || host.slice(-19) === '.flow.microsoft.com') {
    return endpoint;
  }
  return 'https://api.flow.microsoft.com';
}

function flowPath(environment: IFlowEnvironment, flowId: string): string {
  return `${flowApiRoot(environment)}/providers/Microsoft.ProcessSimple/environments/${encodeURIComponent(environment.id)}/flows/${encodeURIComponent(flowId)}`;
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
