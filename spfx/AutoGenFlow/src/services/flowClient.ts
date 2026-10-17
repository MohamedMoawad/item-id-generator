import { AadHttpClient, AadHttpClientFactory } from '@microsoft/sp-http';
import { buildCreateFlowPayload, IListFlowTarget } from './siteFlow';
import {
  FlowConsentError,
  FlowRunState,
  FlowStartError,
  IFlowConnection,
  IFlowEnvironment,
  isConsentFailure,
  readConnections,
  readEnvironments,
  readFlowState
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
  const payload = await sendJson<Parameters<typeof readConnections>[0]>(
    client,
    'GET',
    `${environment.endpoint}/providers/Microsoft.ProcessSimple/environments/${encodeURIComponent(environment.id)}/connections?api-version=${API_VERSION}`
  );
  return readConnections(payload);
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
  await sendJson(
    client,
    'PATCH',
    flowUrl(environment, flowId),
    buildCreateFlowPayload(target, connectionName)
  );
  await startFlow(client, environment, flowId, 'The numbering flow was updated and did not start.');
  return { id: flowId, state: 'Started' };
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
