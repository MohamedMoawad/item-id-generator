export interface IFlowConnection {
  name: string;
  displayName: string;
}

export interface IFlowEnvironment {
  id: string;
  endpoint: string;
  displayName: string;
}

export type FlowRunState = 'Started' | 'Stopped' | 'Missing';

export class FlowStartError extends Error {
  public readonly flowId: string;

  public constructor(flowId: string, message: string) {
    super(message);
    this.name = 'FlowStartError';
    this.flowId = flowId;
    Object.setPrototypeOf(this, FlowStartError.prototype);
  }
}

export class FlowConsentError extends Error {
  public constructor() {
    super('A SharePoint admin approves AutoGen Flow once under Admin center, Advanced, API access. Then turn numbering on.');
    this.name = 'FlowConsentError';
    Object.setPrototypeOf(this, FlowConsentError.prototype);
  }
}

interface IRawEnvironment {
  name?: string;
  properties?: {
    displayName?: string;
    isDefault?: boolean;
    runtimeEndpoints?: { [name: string]: string };
  };
}

interface IRawConnection {
  name?: string;
  properties?: {
    displayName?: string;
    apiId?: string;
    statuses?: Array<{ status?: string }>;
  };
}

export function chooseEnvironment(payload: { value?: IRawEnvironment[] }): IFlowEnvironment | undefined {
  const rows = payload.value || [];
  const preferred = rows.filter((row) => row.properties && row.properties.isDefault)[0] || rows[0];
  if (!preferred || !preferred.name) {
    return undefined;
  }
  const endpoints = preferred.properties && preferred.properties.runtimeEndpoints;
  const endpoint = endpoints && endpoints['microsoft.Flow'] ? endpoints['microsoft.Flow'] : 'https://api.flow.microsoft.com';
  return {
    id: preferred.name,
    endpoint: endpoint.replace(/\/+$/, ''),
    displayName: (preferred.properties && preferred.properties.displayName) || 'Default'
  };
}

export function readConnections(payload: { value?: IRawConnection[] }): IFlowConnection[] {
  return (payload.value || [])
    .filter((row) => {
      const api = ((row.properties && row.properties.apiId) || '').toLowerCase();
      const name = (row.name || '').toLowerCase();
      const connected = !row.properties || !row.properties.statuses || row.properties.statuses.length === 0
        || row.properties.statuses.some((status) => (status.status || '').toLowerCase() === 'connected');
      return connected && (api.indexOf('shared_sharepointonline') >= 0 || name.indexOf('shared-sharepointonl') === 0);
    })
    .map((row) => ({
      name: row.name || '',
      displayName: (row.properties && row.properties.displayName) || row.name || 'SharePoint'
    }))
    .filter((row) => row.name.length > 0);
}

export function readFlowState(payload: { name?: string; properties?: { state?: string } } | undefined): FlowRunState {
  const state = payload && payload.properties && payload.properties.state;
  if (state === 'Started' || state === 'Stopped') {
    return state;
  }
  return 'Missing';
}

export function isConsentFailure(status: number, body: string): boolean {
  if (status === 401 || status === 403) {
    return true;
  }
  const text = body.toLowerCase();
  return text.indexOf('consent') >= 0
    || text.indexOf('aadsts65002') >= 0
    || text.indexOf('aadsts') >= 0
    || text.indexOf('approval') >= 0;
}
