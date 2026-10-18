import { AadHttpClient, AadHttpClientFactory, AadTokenProviderFactory } from '@microsoft/sp-http';
import { buildCreateFlowPayload, buildDataverseWorkflowBody, buildShellFlowPayload, IListFlowTarget } from './siteFlow';
import {
  connectionListUrls,
  connectionPortalUrl,
  flowApiHosts,
  flowCreateUrl,
  flowHomeUrl,
  publicFlowRedirect,
  FlowConsentError,
  FlowRunState,
  FlowStartError,
  ICreatedConnection,
  IFlowConnection,
  IFlowEnvironment,
  IRawListedFlow,
  chooseEnvironment,
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

const dataverseHeaders: { [name: string]: string } = {
  Accept: 'application/json',
  'Content-Type': 'application/json',
  'OData-MaxVersion': '4.0',
  'OData-Version': '4.0',
  Prefer: 'return=representation'
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
  const hosts = flowApiHosts(environment);
  for (let hostIndex = 0; hostIndex < hosts.length; hostIndex += 1) {
    const host = hosts[hostIndex];
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
  active: boolean,
  tokens?: AadTokenProviderFactory
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
  const names = ensured.connection ? [ensured.connection.name] : [''];
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
        const remote = await redeployRemoteFlow(factory, environment, current.flowId, target, connectionName, tokens);
        return startedFlow(environment, remote.id, connectionName, ensured.consentUrl);
      }
      const created = await createRemoteFlow(factory, environment, target, connectionName, tokens);
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
  connectionName: string,
  tokens?: AadTokenProviderFactory
): Promise<IRemoteFlow> {
  const client = await openClient(factory);
  const token = await readFlowToken(tokens);
  const located = await environmentForWrite(client, token, environment);
  const hosts = flowApiHosts(located);
  const full = buildCreateFlowPayload(target, connectionName);
  const shell = buildShellFlowPayload(target.listTitle);
  let lastError = '';
  try {
    const dataverse = await createDataverseFlow(factory, located, target, connectionName);
    if (dataverse) {
      return dataverse;
    }
  } catch (error) {
    if (error instanceof FlowStartError) {
      throw error;
    }
    lastError = error instanceof Error ? error.message : lastError;
  }
  for (let hostIndex = 0; hostIndex < hosts.length; hostIndex += 1) {
    const host = hosts[hostIndex];
    const created = await createOnHost(client, token, environment, host, full, shell);
    if (!created.id) {
      lastError = preferFlowError(lastError, created.error);
      continue;
    }
    try {
      await startFlow(client, environment, created.id, 'The numbering flow was created and did not start.', host, token);
      return { id: created.id, state: 'Started' };
    } catch (error) {
      if (error instanceof FlowStartError) {
        throw error;
      }
      lastError = preferFlowError(lastError, error instanceof Error ? error.message : lastError);
    }
  }
  throw new Error(publicFlowError(environment.id, lastError));
}

export async function redeployRemoteFlow(
  factory: AadHttpClientFactory,
  environment: IFlowEnvironment,
  flowId: string,
  target: IListFlowTarget,
  connectionName: string,
  tokens?: AadTokenProviderFactory
): Promise<IRemoteFlow> {
  const client = await openClient(factory);
  const token = await readFlowToken(tokens);
  const located = await environmentForWrite(client, token, environment);
  const dataverse = await redeployDataverseFlow(factory, located, flowId, target, connectionName);
  if (dataverse) {
    return dataverse;
  }
  const host = flowApiHosts(located)[0] || 'https://api.flow.microsoft.com';
  try {
    await sendJson(client, 'POST', `${flowPath(environment, flowId, host)}/stop?api-version=${API_VERSION}`, undefined, token);
  } catch (error) {
    if (error instanceof FlowConsentError) {
      throw error;
    }
  }
  await sendJson(
    client,
    'PATCH',
    flowUrlOn(environment, flowId, host),
    buildCreateFlowPayload(target, connectionName),
    token
  );
  await startFlow(client, environment, flowId, 'The numbering flow was updated and did not start.', host, token);
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

async function startFlow(client: AadHttpClient, environment: IFlowEnvironment, flowId: string, fallback: string, host?: string, token?: string): Promise<void> {
  const preferred = host || flowApiHosts(environment)[0] || 'https://api.flow.microsoft.com';
  const hosts = [preferred].concat(flowApiHosts(environment).filter((item) => item !== preferred));
  let lastError = '';
  for (let index = 0; index < hosts.length; index += 1) {
    try {
      await sendJson(client, 'POST', `${flowPath(environment, flowId, hosts[index])}/start?api-version=${API_VERSION}`, undefined, token);
      return;
    } catch (error) {
      if (error instanceof FlowConsentError) {
        throw error;
      }
      lastError = preferFlowError(lastError, error instanceof Error ? error.message : '');
    }
  }
  throw new FlowStartError(flowId, lastError || fallback);
}

async function openConnectionClients(factory: AadHttpClientFactory): Promise<{ powerApps?: AadHttpClient; ordered: AadHttpClient[] }> {
  const flow = await openClient(factory);
  return { ordered: [flow] };
}

async function openClient(factory: AadHttpClientFactory): Promise<AadHttpClient> {
  return factory.getClient(FLOW_RESOURCE);
}

async function environmentForWrite(client: AadHttpClient, token: string, environment: IFlowEnvironment): Promise<IFlowEnvironment> {
  try {
    const detail = await sendJson<NonNullable<Parameters<typeof chooseEnvironment>[0]['value']>[number]>(
      client,
      'GET',
      `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/${encodeURIComponent(environment.id)}?api-version=${API_VERSION}`,
      undefined,
      token
    );
    const refreshed = chooseEnvironment({ value: [detail] });
    if (refreshed && refreshed.id === environment.id && refreshed.location) {
      return refreshed;
    }
  } catch (error) {
    if (error instanceof FlowConsentError) {
      throw error;
    }
  }
  return environment;
}

async function readFlowToken(tokens?: AadTokenProviderFactory): Promise<string> {
  if (!tokens) {
    return '';
  }
  try {
    const provider = await tokens.getTokenProvider();
    return await provider.getToken(FLOW_RESOURCE);
  } catch {
    return '';
  }
}

interface IFlowWriteResult {
  id: string;
  error: string;
}

interface IFlowWriteResponse {
  name?: string;
  id?: string;
}

async function createDataverseFlow(
  factory: AadHttpClientFactory,
  environment: IFlowEnvironment,
  target: IListFlowTarget,
  connectionName: string
): Promise<IRemoteFlow | undefined> {
  const client = await openDataverseClient(factory, environment.instanceUrl);
  if (!client) {
    return undefined;
  }
  const response = await client.fetch(`${environment.instanceUrl}/api/data/v9.2/workflows`, AadHttpClient.configurations.v1, {
    method: 'POST',
    headers: dataverseHeaders,
    body: JSON.stringify(buildDataverseWorkflowBody(target, connectionName))
  });
  const text = await response.text();
  if (!response.ok) {
    throw new FlowRequestError(response.status, `${environment.instanceUrl} returned HTTP ${response.status}. ${cleanFlowError(shorten(text, response.status))}`);
  }
  const body = text.trim() ? parseJson<{ workflowid?: string }>(text) : {};
  const id = body.workflowid || workflowIdFromHeader(response.headers.get('OData-EntityId') || '');
  if (!id) {
    throw new Error('Dataverse did not return the new numbering flow.');
  }
  await activateDataverseFlow(client, environment.instanceUrl, id);
  return { id, state: 'Started' };
}

async function redeployDataverseFlow(
  factory: AadHttpClientFactory,
  environment: IFlowEnvironment,
  flowId: string,
  target: IListFlowTarget,
  connectionName: string
): Promise<IRemoteFlow | undefined> {
  const client = await openDataverseClient(factory, environment.instanceUrl);
  if (!client) {
    return undefined;
  }
  const row = buildDataverseWorkflowBody(target, connectionName);
  const url = `${environment.instanceUrl}/api/data/v9.2/workflows(${flowId})`;
  try {
    await dataverseSend(client, 'PATCH', url, { statecode: 0 });
    await dataverseSend(client, 'PATCH', url, { clientdata: row.clientdata, name: row.name });
  } catch (error) {
    if (error instanceof FlowConsentError) {
      throw error;
    }
    return undefined;
  }
  await activateDataverseFlow(client, environment.instanceUrl, flowId);
  return { id: flowId, state: 'Started' };
}

async function openDataverseClient(factory: AadHttpClientFactory, orgUrl: string): Promise<AadHttpClient | undefined> {
  if (!orgUrl) {
    return undefined;
  }
  try {
    return await factory.getClient(orgUrl);
  } catch {
    return undefined;
  }
}

async function activateDataverseFlow(client: AadHttpClient, orgUrl: string, flowId: string): Promise<void> {
  try {
    await dataverseSend(client, 'PATCH', `${orgUrl}/api/data/v9.2/workflows(${flowId})`, { statecode: 1 });
  } catch (error) {
    const message = error instanceof Error ? error.message : 'The numbering flow was created and did not start.';
    throw new FlowStartError(flowId, message);
  }
}

async function dataverseSend(client: AadHttpClient, method: string, url: string, body: object): Promise<void> {
  const response = await client.fetch(url, AadHttpClient.configurations.v1, {
    method,
    headers: dataverseHeaders,
    body: JSON.stringify(body)
  });
  if (response.ok) {
    return;
  }
  const text = await response.text();
  throw new FlowRequestError(response.status, `${hostOf(url)} returned HTTP ${response.status}. ${cleanFlowError(shorten(text, response.status))}`);
}

function workflowIdFromHeader(value: string): string {
  const match = /workflows\(([0-9a-f-]{36})\)/i.exec(value);
  return match ? match[1] : '';
}

async function createOnHost(
  client: AadHttpClient,
  token: string,
  environment: IFlowEnvironment,
  host: string,
  full: object,
  shell: object
): Promise<IFlowWriteResult> {
  const collection = flowCreateUrl(host, environment.id);
  const shellResult = await writeFlow(client, token, 'POST', collection, shell);
  if (shellResult.id) {
    const patched = await writeFlow(client, token, 'PATCH', flowUrlOn(environment, shellResult.id, host), full);
    if (!patched.error) {
      return { id: shellResult.id, error: '' };
    }
    await deleteFlowQuiet(client, token, environment, shellResult.id, host);
    shellResult.error = preferFlowError(shellResult.error, patched.error);
  }
  const flowId = newFlowId();
  const put = await writeFlow(client, token, 'PUT', flowUrlOn(environment, flowId, host), full);
  if (!put.error) {
    return { id: put.id || flowId, error: '' };
  }
  const posted = await writeFlow(client, token, 'POST', collection, full);
  if (!posted.error && posted.id) {
    return { id: posted.id, error: '' };
  }
  return {
    id: '',
    error: preferFlowError(preferFlowError(shellResult.error, put.error), posted.error)
  };
}

async function writeFlow(client: AadHttpClient, token: string, method: string, url: string, body: object): Promise<IFlowWriteResult> {
  try {
    const payload = await sendJson<IFlowWriteResponse>(client, method, url, body, token);
    return { id: createdFlowId(payload), error: '' };
  } catch (error) {
    if (error instanceof FlowConsentError) {
      throw error;
    }
    return { id: '', error: error instanceof Error ? error.message : '' };
  }
}

async function deleteFlowQuiet(client: AadHttpClient, token: string, environment: IFlowEnvironment, flowId: string, host: string): Promise<void> {
  try {
    await sendJson(client, 'DELETE', flowUrlOn(environment, flowId, host), undefined, token);
  } catch (error) {
    if (error instanceof FlowConsentError) {
      throw error;
    }
  }
}

function createdFlowId(payload: IFlowWriteResponse): string {
  if (payload.name) {
    return payload.name;
  }
  const match = /\/flows\/([^/?]+)/i.exec(payload.id || '');
  return match ? decodeURIComponent(match[1]) : '';
}

function newFlowId(): string {
  const source = typeof crypto !== 'undefined' ? crypto : undefined;
  if (source && typeof source.randomUUID === 'function') {
    return source.randomUUID();
  }
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (char) => {
    const random = Math.floor(Math.random() * 16);
    const value = char === 'x' ? random : (random % 4) + 8;
    return value.toString(16);
  });
}

async function sendJson<T>(client: AadHttpClient, method: string, url: string, body?: object, token?: string): Promise<T> {
  if (token) {
    return sendWithToken<T>(token, method, url, body);
  }
  const response = await client.fetch(url, AadHttpClient.configurations.v1, {
    method,
    headers: jsonHeaders,
    body: body === undefined ? undefined : JSON.stringify(body),
    redirect: 'error'
  });
  const text = await response.text();
  if (!response.ok) {
    throw new FlowRequestError(response.status, `${hostOf(url)} returned HTTP ${response.status}. ${cleanFlowError(shorten(text, response.status))}`);
  }
  return parseJson<T>(text);
}

async function sendWithToken<T>(token: string, method: string, url: string, body?: object, hop = 0): Promise<T> {
  const response = await fetch(url, {
    method,
    headers: {
      Accept: 'application/json',
      'Content-Type': 'application/json',
      Authorization: `Bearer ${token}`
    },
    body: body === undefined ? undefined : JSON.stringify(body),
    redirect: 'manual'
  });
  if (response.type === 'opaqueredirect' || response.status === 0) {
    throw new FlowRequestError(0, `${hostOf(url)} redirected away from the public Flow service.`);
  }
  if (response.status >= 300 && response.status < 400) {
    const next = hop < 1 ? publicFlowRedirect(url, response.headers.get('Location') || '') : '';
    if (next) {
      return sendWithToken<T>(token, method, next, body, hop + 1);
    }
    throw new FlowRequestError(response.status, `${hostOf(url)} returned HTTP ${response.status}.`);
  }
  const text = await response.text();
  if (!response.ok) {
    throw new FlowRequestError(response.status, `${hostOf(url)} returned HTTP ${response.status}. ${cleanFlowError(shorten(text, response.status))}`);
  }
  return parseJson<T>(text);
}

function parseJson<T>(text: string): T {
  if (!text.trim()) {
    return {} as T;
  }
  try {
    return JSON.parse(text) as T;
  } catch {
    throw new Error('Power Automate returned a response that was not JSON.');
  }
}

function preferFlowError(current: string, next: string): string {
  if (!next) {
    return current;
  }
  if (flowErrorRank(next) > flowErrorRank(current)) {
    return next;
  }
  return current;
}

function flowErrorRank(message: string): number {
  if (!message) {
    return 0;
  }
  if (message.indexOf('No HTTP resource was found') >= 0) {
    return 1;
  }
  if (message.indexOf('Failed to fetch') >= 0 || message.indexOf('NetworkError') >= 0) {
    return 2;
  }
  return 3;
}

function publicFlowError(environmentId: string, message: string): string {
  const detail = cleanFlowError(message);
  if (detail) {
    return detail;
  }
  return `The numbering flow was not created in ${flowHomeUrl(environmentId)}.`;
}

function cleanFlowError(message: string): string {
  return message
    .replace(/https:\/\/10\.[0-9.]+(?::[0-9]+)?[^'"\s]*/g, '')
    .replace(/The caller does not have permission to perform [A-Za-z]+ on api '[^']*'\.?/gi, '')
    .replace(/\s+/g, ' ')
    .replace(/\s+\./g, '.')
    .trim();
}

function hostOf(url: string): string {
  const match = /^https:\/\/[^/]+/i.exec(url);
  return match ? match[0] : 'https://api.flow.microsoft.com';
}

function flowPath(environment: IFlowEnvironment, flowId: string, host?: string): string {
  const root = host || flowApiHosts(environment)[0] || 'https://api.flow.microsoft.com';
  return `${root}/providers/Microsoft.ProcessSimple/environments/${encodeURIComponent(environment.id)}/flows/${encodeURIComponent(flowId)}`;
}

function flowUrl(environment: IFlowEnvironment, flowId: string): string {
  return `${flowPath(environment, flowId)}?api-version=${API_VERSION}`;
}

function flowUrlOn(environment: IFlowEnvironment, flowId: string, host: string): string {
  return `${flowPath(environment, flowId, host)}?api-version=${API_VERSION}`;
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
