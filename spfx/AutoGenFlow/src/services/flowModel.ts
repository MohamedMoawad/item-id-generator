export interface IFlowConnection {
  name: string;
  displayName: string;
}

export interface IFlowEnvironment {
  id: string;
  endpoint: string;
  powerAppsEndpoint: string;
  displayName: string;
  isDefault: boolean;
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
    environment?: { name?: string };
  };
}

export function readEnvironments(payload: { value?: IRawEnvironment[] }): IFlowEnvironment[] {
  const rows = (payload.value || [])
    .map(toEnvironment)
    .filter((row): row is IFlowEnvironment => row !== undefined);
  rows.sort((left, right) => {
    if (left.isDefault !== right.isDefault) {
      return left.isDefault ? -1 : 1;
    }
    return left.displayName.localeCompare(right.displayName);
  });
  return rows;
}

export function chooseEnvironment(payload: { value?: IRawEnvironment[] }): IFlowEnvironment | undefined {
  return readEnvironments(payload)[0];
}

function toEnvironment(row: IRawEnvironment): IFlowEnvironment | undefined {
  if (!row.name) {
    return undefined;
  }
  const endpoints = row.properties && row.properties.runtimeEndpoints;
  const endpoint = endpoints && endpoints['microsoft.Flow'] ? endpoints['microsoft.Flow'] : 'https://api.flow.microsoft.com';
  const powerApps = endpoints && endpoints['microsoft.PowerApps'] ? endpoints['microsoft.PowerApps'] : 'https://api.powerapps.com';
  return {
    id: row.name,
    endpoint: endpoint.replace(/\/+$/, ''),
    powerAppsEndpoint: powerApps.replace(/\/+$/, ''),
    displayName: (row.properties && row.properties.displayName) || 'Default',
    isDefault: !!(row.properties && row.properties.isDefault)
  };
}

const API_VERSION = '2016-11-01';

export function connectionListUrls(environment: IFlowEnvironment): string[] {
  const hosts = [
    'https://api.flow.microsoft.com',
    environment.endpoint,
    environment.powerAppsEndpoint
  ];
  const urls: string[] = [];
  hosts.forEach((host) => {
    pushUrl(urls, sharePointConnectionUrl(host, environment.id));
  });
  pushUrl(urls, `${environment.endpoint}/providers/Microsoft.ProcessSimple/environments/${encodeURIComponent(environment.id)}/connections?api-version=${API_VERSION}`);
  pushUrl(urls, `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/${encodeURIComponent(environment.id)}/connections?api-version=${API_VERSION}`);
  return urls;
}

function sharePointConnectionUrl(host: string, environmentId: string): string {
  const filter = encodeURIComponent(`environment eq '${environmentId.replace(/'/g, "''")}'`);
  return `${host.replace(/\/+$/, '')}/providers/Microsoft.PowerApps/apis/shared_sharepointonline/connections?api-version=${API_VERSION}&$filter=${filter}`;
}

function pushUrl(urls: string[], url: string): void {
  if (urls.indexOf(url) === -1) {
    urls.push(url);
  }
}

export function readConnections(payload: { value?: IRawConnection[] }, sharePointScoped = false, environmentId = ''): IFlowConnection[] {
  return (payload.value || [])
    .filter((row) => {
      const api = ((row.properties && row.properties.apiId) || '').toLowerCase();
      const name = (row.name || '').toLowerCase();
      const rowEnvironment = row.properties && row.properties.environment && row.properties.environment.name;
      const connected = !row.properties || !row.properties.statuses || row.properties.statuses.length === 0
        || row.properties.statuses.some((status) => (status.status || '').toLowerCase() === 'connected');
      const sharePoint = sharePointScoped
        || api.indexOf('shared_sharepointonline') >= 0
        || api.indexOf('sharepoint') >= 0
        || name.indexOf('shared-sharepointonl') === 0;
      const otherApi = api.length > 0 && api.indexOf('sharepoint') < 0;
      const sameEnvironment = !environmentId || !rowEnvironment || rowEnvironment === environmentId;
      return connected && sharePoint && !otherApi && sameEnvironment;
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
