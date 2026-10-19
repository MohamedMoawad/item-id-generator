export interface IFlowConnection {
  name: string;
  displayName: string;
  accountName: string;
}

export interface IFlowEnvironment {
  id: string;
  endpoint: string;
  powerAppsEndpoint: string;
  displayName: string;
  isDefault: boolean;
  location: string;
  instanceUrl: string;
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
  public constructor(message?: string) {
    super(message || 'Open Admin center, Advanced, API access. Microsoft Flow Service needs Flows.Read.All and Flows.Manage.All approved. Then turn numbering on.');
    this.name = 'FlowConsentError';
    Object.setPrototypeOf(this, FlowConsentError.prototype);
  }
}

interface IRawEnvironment {
  name?: string;
  location?: string;
  properties?: {
    displayName?: string;
    isDefault?: boolean;
    location?: string;
    region?: string;
    azureRegion?: string;
    azureRegionHint?: string;
    runtimeEndpoints?: { [name: string]: string };
    linkedEnvironmentMetadata?: { instanceUrl?: string; instanceApiUrl?: string };
  };
}

interface IRawConnection {
  name?: string;
  id?: string;
  properties?: {
    displayName?: string;
    apiId?: string;
    accountName?: string;
    api?: { name?: string };
    statuses?: Array<{ status?: string }>;
    environment?: string | { name?: string; id?: string };
    createdBy?: { displayName?: string; email?: string };
    consentInfo?: { redirectUrl?: string; firstPartyLoginUri?: string; consentLink?: string };
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
  const flowRuntime = endpoints && endpoints['microsoft.Flow'];
  const powerAppsRuntime = endpoints && endpoints['microsoft.PowerApps'];
  return {
    id: row.name,
    endpoint: publicServiceUrl(flowRuntime, 'https://api.flow.microsoft.com', '.flow.microsoft.com'),
    powerAppsEndpoint: publicServiceUrl(powerAppsRuntime, 'https://api.powerapps.com', '.powerapps.com'),
    displayName: (row.properties && row.properties.displayName) || 'Default',
    isDefault: !!(row.properties && row.properties.isDefault),
    location: canonicalFlowGeo(row),
    instanceUrl: dataverseOrgUrl(row)
  };
}

const FLOW_GEOS = [
  'unitedstates', 'europe', 'asia', 'australia', 'india', 'japan', 'canada', 'unitedkingdom',
  'france', 'unitedarabemirates', 'germany', 'switzerland', 'norway', 'korea', 'southafrica',
  'sweden', 'singapore', 'poland', 'italy', 'israel', 'spain', 'mexico', 'southamerica'
];

const REGION_GEOS: { [name: string]: string } = {
  uaenorth: 'unitedarabemirates',
  uaecentral: 'unitedarabemirates',
  qatarcentral: 'unitedarabemirates',
  qatar: 'unitedarabemirates',
  westeurope: 'europe',
  northeurope: 'europe',
  francecentral: 'france',
  francesouth: 'france',
  germanywestcentral: 'germany',
  germanynorth: 'germany',
  norwayeast: 'norway',
  norwaywest: 'norway',
  uksouth: 'unitedkingdom',
  ukwest: 'unitedkingdom',
  switzerlandnorth: 'switzerland',
  switzerlandwest: 'switzerland',
  swedencentral: 'sweden',
  southafricanorth: 'southafrica',
  southafricawest: 'southafrica',
  westus: 'unitedstates',
  westus2: 'unitedstates',
  westus3: 'unitedstates',
  eastus: 'unitedstates',
  eastus2: 'unitedstates',
  centralus: 'unitedstates',
  northcentralus: 'unitedstates',
  southcentralus: 'unitedstates',
  westcentralus: 'unitedstates',
  canadaeast: 'canada',
  canadacentral: 'canada',
  brazilsouth: 'southamerica',
  australiaeast: 'australia',
  australiasoutheast: 'australia',
  australiacentral: 'australia',
  eastasia: 'asia',
  southeastasia: 'asia',
  japaneast: 'japan',
  japanwest: 'japan',
  centralindia: 'india',
  southindia: 'india',
  westindia: 'india',
  koreacentral: 'korea',
  koreasouth: 'korea',
  polandcentral: 'poland',
  italynorth: 'italy',
  israelcentral: 'israel',
  spaincentral: 'spain',
  mexicocentral: 'mexico'
};

function canonicalFlowGeo(row: IRawEnvironment): string {
  const properties = row.properties;
  const metadata = properties && properties.linkedEnvironmentMetadata;
  const datacenter = [
    (properties && properties.azureRegionHint) || '',
    (properties && properties.azureRegion) || '',
    (properties && properties.region) || '',
    (metadata && metadata.instanceUrl) || '',
    (metadata && metadata.instanceApiUrl) || ''
  ];
  for (let index = 0; index < datacenter.length; index += 1) {
    const geo = geoFromHint(datacenter[index]) || geoFromInstance(datacenter[index]);
    if (geo) {
      return geo;
    }
  }
  const hints = [
    row.location || '',
    (properties && properties.location) || ''
  ];
  for (let index = 0; index < hints.length; index += 1) {
    const geo = geoFromHint(hints[index]) || geoFromInstance(hints[index]);
    if (geo) {
      return geo;
    }
  }
  const endpoints = properties && properties.runtimeEndpoints;
  if (!endpoints) {
    return '';
  }
  const keys = Object.keys(endpoints);
  for (let index = 0; index < keys.length; index += 1) {
    const geo = geoFromFlowHost(publicFlowHost(endpoints[keys[index]]));
    if (geo) {
      return geo;
    }
  }
  return '';
}

export function geoFromHint(value: string): string {
  const key = value.trim().toLowerCase().replace(/[^a-z0-9]/g, '');
  if (!key) {
    return '';
  }
  if (FLOW_GEOS.indexOf(key) >= 0) {
    return key;
  }
  return REGION_GEOS[key] || '';
}

function geoFromInstance(value: string): string {
  const match = /crm(\d*)\.dynamics\.com/i.exec(value);
  if (!match) {
    return '';
  }
  const geos: { [name: string]: string } = {
    '': 'unitedstates',
    '2': 'southamerica',
    '3': 'canada',
    '4': 'europe',
    '5': 'asia',
    '6': 'australia',
    '7': 'japan',
    '8': 'india',
    '11': 'unitedkingdom',
    '12': 'france',
    '14': 'southafrica',
    '15': 'unitedarabemirates',
    '16': 'germany',
    '17': 'switzerland',
    '19': 'norway',
    '20': 'singapore',
    '21': 'korea',
    '22': 'sweden'
  };
  return geos[match[1]] || '';
}

export function dataverseOrgUrl(row: IRawEnvironment): string {
  const metadata = row.properties && row.properties.linkedEnvironmentMetadata;
  const raw = ((metadata && (metadata.instanceUrl || metadata.instanceApiUrl)) || '').trim().replace(/\/+$/, '');
  if (!raw) {
    return '';
  }
  let host = raw;
  const withProtocol = /^https:\/\//i.test(host) ? host : '';
  if (!withProtocol) {
    return '';
  }
  host = withProtocol.replace(/\.api\.crm/i, '.crm').toLowerCase();
  if (!/^https:\/\/[a-z0-9-]+\.crm\d*\.dynamics\.com$/i.test(host)) {
    return '';
  }
  return host;
}

function geoFromFlowHost(host: string): string {
  const match = /^https:\/\/([a-z0-9]+)\.api\.flow\.microsoft\.com$/i.exec(host);
  if (!match || match[1] === 'api') {
    return '';
  }
  return geoFromHint(match[1]);
}

function publicServiceUrl(value: string | undefined, fallback: string, suffix: string): string {
  if (!value) {
    return fallback;
  }
  const trimmed = value.trim().replace(/\/+$/, '');
  const match = /^https:\/\/([^/:?#]+)(?::\d+)?$/i.exec(trimmed);
  if (!match) {
    return fallback;
  }
  const host = match[1].toLowerCase();
  if (isPrivateHost(host) || (host !== suffix.slice(1) && host.indexOf(suffix) !== host.length - suffix.length)) {
    return fallback;
  }
  return trimmed;
}

function isPrivateHost(host: string): boolean {
  if (host === 'localhost' || host.slice(-6) === '.local') {
    return true;
  }
  const parts = host.split('.');
  if (parts.length !== 4) {
    return false;
  }
  const numbers = parts.map((part) => (/^\d+$/.test(part) ? Number(part) : -1));
  if (numbers.some((part) => part < 0 || part > 255)) {
    return false;
  }
  return numbers[0] === 10
    || numbers[0] === 127
    || (numbers[0] === 192 && numbers[1] === 168)
    || (numbers[0] === 172 && numbers[1] >= 16 && numbers[1] <= 31)
    || (numbers[0] === 169 && numbers[1] === 254);
}

const API_VERSION = '2016-11-01';

export function connectionListUrls(environment: IFlowEnvironment): string[] {
  const powerApps = environment.powerAppsEndpoint.replace(/\/+$/, '');
  const flow = environment.endpoint.replace(/\/+$/, '');
  const filter = encodeURIComponent(`environment eq '${environment.id.replace(/'/g, "''")}'`);
  const version = `api-version=${API_VERSION}`;
  const id = encodeURIComponent(environment.id);
  const hosts = [powerApps, 'https://api.powerapps.com', flow, 'https://api.flow.microsoft.com'];
  const urls: string[] = [];
  hosts.forEach((host) => {
    pushUrl(urls, `${host}/providers/Microsoft.PowerApps/connections?${version}&$filter=${filter}`);
    pushUrl(urls, `${host}/providers/Microsoft.PowerApps/apis/shared_sharepointonline/connections?${version}&$filter=${filter}`);
  });
  pushUrl(urls, `${flow}/providers/Microsoft.ProcessSimple/environments/${id}/connections?${version}`);
  pushUrl(urls, `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/${id}/connections?${version}`);
  return urls;
}

function pushUrl(urls: string[], url: string): void {
  if (urls.indexOf(url) === -1) {
    urls.push(url);
  }
}

export function sharePointConnectionCreateUrls(environment: IFlowEnvironment): string[] {
  const filter = encodeURIComponent(`environment eq '${environment.id.replace(/'/g, "''")}'`);
  const path = `/providers/Microsoft.PowerApps/apis/shared_sharepointonline/connections?api-version=${API_VERSION}&$filter=${filter}`;
  const hosts = [
    'https://api.flow.microsoft.com',
    environment.endpoint.replace(/\/+$/, ''),
    environment.powerAppsEndpoint.replace(/\/+$/, ''),
    'https://api.powerapps.com'
  ];
  const urls: string[] = [];
  hosts.forEach((host) => pushUrl(urls, `${host}${path}`));
  return urls;
}

export function sharePointConnectionPutUrl(environment: IFlowEnvironment, connectionName: string): string {
  return sharePointConnectionCreateUrls(environment)[0].replace(
    '/connections?',
    `/connections/${encodeURIComponent(connectionName)}?`
  );
}

export function sharePointConnectionBody(environmentId: string): { properties: { displayName: string; apiId: string; connectionParameters: object; environment: { name: string; id: string } } } {
  return {
    properties: {
      displayName: 'SharePoint',
      apiId: '/providers/Microsoft.PowerApps/apis/shared_sharepointonline',
      connectionParameters: {},
      environment: {
        name: environmentId,
        id: `/providers/Microsoft.PowerApps/environments/${environmentId}`
      }
    }
  };
}

export interface ICreatedConnection {
  name: string;
  displayName: string;
  accountName: string;
  consentUrl: string;
  connected: boolean;
}

export function readCreatedConnection(payload: { value?: IRawConnection[]; name?: string; id?: string; properties?: IRawConnection['properties'] }): ICreatedConnection | undefined {
  const row = payload.value && payload.value.length > 0 ? payload.value[0] : payload;
  const name = connectionName(row);
  if (!name || !isSharePointConnection(row, true)) {
    return undefined;
  }
  const connected = connectionIsReady(row);
  return {
    name,
    displayName: connectionUser(row),
    accountName: connectionAccount(row),
    consentUrl: connected ? '' : consentUrlOf(row),
    connected
  };
}

export function connectionPortalUrl(environmentId: string): string {
  const environment = encodeURIComponent(environmentId.trim());
  return `https://make.powerautomate.com/environments/${environment}/connections`;
}

function consentUrlOf(row: IRawConnection): string {
  const info = row.properties && row.properties.consentInfo;
  if (!info) {
    return '';
  }
  return info.redirectUrl || info.firstPartyLoginUri || info.consentLink || '';
}

export function readConnections(payload: { value?: IRawConnection[] }, sharePointScoped = false, environmentId = ''): IFlowConnection[] {
  const rows = (payload.value || [])
    .filter((row) => isSharePointConnection(row, sharePointScoped) && inEnvironment(row, environmentId))
    .map((row) => ({
      name: connectionName(row),
      displayName: connectionUser(row),
      accountName: connectionAccount(row),
      connected: connectionIsReady(row)
    }))
    .filter((row) => row.name.length > 0);
  rows.sort((left, right) => {
    if (left.connected === right.connected) {
      return left.displayName.localeCompare(right.displayName);
    }
    return left.connected ? -1 : 1;
  });
  return rows.map((row) => ({ name: row.name, displayName: row.displayName, accountName: row.accountName }));
}

export function chooseConnection(connections: IFlowConnection[], account: string, savedName = ''): IFlowConnection | undefined {
  const needle = normalizeAccount(account);
  if (needle) {
    const mine = connections.filter((item) => normalizeAccount(item.accountName) === needle || normalizeAccount(item.displayName) === needle);
    if (mine.length > 0) {
      return mine[0];
    }
  }
  const saved = connections.filter((item) => item.name === savedName)[0];
  return saved || connections[0];
}

export function flowPortalUrl(environmentId: string, flowId: string): string {
  const environment = encodeURIComponent(environmentId.trim());
  const flow = encodeURIComponent(flowId.trim());
  return `https://make.powerautomate.com/environments/${environment}/flows/${flow}/details`;
}

export function flowHomeUrl(environmentId: string): string {
  const environment = encodeURIComponent(environmentId.trim());
  return `https://make.powerautomate.com/environments/${environment}/home`;
}

export function flowListUrl(environmentId: string): string {
  const environment = encodeURIComponent(environmentId.trim());
  return `https://make.powerautomate.com/environments/${environment}/flows`;
}

export function flowApiHosts(environment: { location?: string; endpoint?: string }): string[] {
  const hosts: string[] = [];
  const geo = geoFromHint(environment.location || '');
  if (geo === 'europe') {
    addFlowHost(hosts, 'https://emea.api.flow.microsoft.com');
    addFlowHost(hosts, 'https://api.flow.microsoft.com');
    return hosts;
  }
  if (geo) {
    addFlowHost(hosts, `https://${geo}.api.flow.microsoft.com`);
    addFlowHost(hosts, 'https://api.flow.microsoft.com');
    return hosts;
  }
  const endpoint = (environment.endpoint || '').toLowerCase();
  if (endpoint.indexOf('europe.api.flow.microsoft.com') >= 0) {
    addFlowHost(hosts, 'https://emea.api.flow.microsoft.com');
    addFlowHost(hosts, 'https://api.flow.microsoft.com');
    return hosts;
  }
  addFlowHost(hosts, environment.endpoint || '');
  addFlowHost(hosts, 'https://api.flow.microsoft.com');
  return hosts;
}

export function flowGeoCandidates(environment: { location?: string; endpoint?: string }): string[] {
  const hosts = flowApiHosts(environment).filter((host) => host !== 'https://api.flow.microsoft.com');
  FLOW_GEOS.forEach((geo) => addFlowHost(hosts, `https://${geo}.api.flow.microsoft.com`));
  return hosts;
}

export function flowCreateUrl(host: string, environmentId: string): string {
  const id = encodeURIComponent(environmentId.trim());
  return `${host}/providers/Microsoft.ProcessSimple/environments/${id}/flows?api-version=2016-11-01`;
}

export function flowPutUrl(host: string, environmentId: string, flowId: string): string {
  const id = encodeURIComponent(environmentId.trim());
  const flow = encodeURIComponent(flowId.trim());
  return `${host}/providers/Microsoft.ProcessSimple/environments/${id}/flows/${flow}?api-version=2016-11-01`;
}

export function publicFlowRedirect(fromUrl: string, locationHeader: string): string {
  const raw = locationHeader.trim();
  if (!raw) {
    return '';
  }
  let absolute = raw;
  if (raw.charAt(0) === '/') {
    const origin = /^https:\/\/[^/]+/i.exec(fromUrl);
    absolute = origin ? `${origin[0]}${raw}` : '';
  }
  if (!/^https:\/\//i.test(absolute)) {
    return '';
  }
  const origin = /^https:\/\/[^/?#]+/i.exec(absolute);
  const fromOrigin = /^https:\/\/[^/?#]+/i.exec(fromUrl);
  if (!origin || !fromOrigin || origin[0].toLowerCase() !== fromOrigin[0].toLowerCase() || !publicFlowHost(origin[0])) {
    return '';
  }
  return absolute;
}

function addFlowHost(hosts: string[], value: string): void {
  const host = publicFlowHost(value);
  if (host && hosts.indexOf(host) === -1) {
    hosts.push(host);
  }
}

function publicFlowHost(value: string): string {
  const trimmed = value.trim().replace(/\/+$/, '');
  const match = /^https:\/\/([^/:?#]+)$/i.exec(trimmed);
  if (!match) {
    return '';
  }
  const host = match[1].toLowerCase();
  if (isPrivateHost(host) || (host !== 'api.flow.microsoft.com' && host.slice(-23) !== '.api.flow.microsoft.com')) {
    return '';
  }
  return `https://${host}`;
}

export function flowCreateUrls(environment: { id: string; location?: string; endpoint?: string }): string[] {
  return flowApiHosts(environment).map((host) => flowCreateUrl(host, environment.id));
}

export interface IRawListedFlow {
  name?: string;
  properties?: {
    connectionReferences?: { [key: string]: IRawConnectionReference };
  };
}

interface IRawConnectionReference {
  connectionName?: string;
  id?: string;
  connection?: { name?: string };
  api?: { name?: string };
}

export function readConnectionNamesFromFlows(payload: { value?: IRawListedFlow[]; properties?: IRawListedFlow['properties'] }): IFlowConnection[] {
  const listed = payload.value || (payload.properties ? [payload] : []);
  const found: IFlowConnection[] = [];
  const seen: { [name: string]: boolean } = {};
  listed.forEach((flow) => {
    const references = flow.properties && flow.properties.connectionReferences;
    if (!references) {
      return;
    }
    Object.keys(references).forEach((key) => {
      const reference = references[key];
      const name = referenceConnectionName(reference);
      if (!name || seen[name] || !isSharePointReference(key, reference)) {
        return;
      }
      seen[name] = true;
      found.push({ name, displayName: name, accountName: '' });
    });
  });
  return found;
}

function referenceConnectionName(reference: IRawConnectionReference | undefined): string {
  if (!reference) {
    return '';
  }
  if (reference.connectionName) {
    return reference.connectionName;
  }
  return reference.connection && reference.connection.name ? reference.connection.name : '';
}

function isSharePointReference(key: string, reference: IRawConnectionReference): boolean {
  const api = reference.api && reference.api.name ? reference.api.name : '';
  const text = `${key} ${reference.id || ''} ${api}`.toLowerCase();
  return text.indexOf('sharepoint') >= 0;
}

function normalizeAccount(value: string): string {
  const trimmed = value.trim().toLowerCase();
  const parts = trimmed.split('|');
  return parts[parts.length - 1] || '';
}

function connectionName(row: IRawConnection): string {
  if (row.name) {
    return row.name;
  }
  const id = row.id || '';
  const parts = id.split('/');
  return parts[parts.length - 1] || '';
}

function isSharePointConnection(row: IRawConnection, sharePointScoped: boolean): boolean {
  const api = connectionApi(row);
  const name = connectionName(row).toLowerCase();
  const id = (row.id || '').toLowerCase();
  const label = ((row.properties && row.properties.displayName) || '').toLowerCase();
  const sharePoint = api.indexOf('sharepoint') >= 0 || name.indexOf('sharepoint') >= 0 || id.indexOf('sharepoint') >= 0 || label.indexOf('sharepoint') >= 0;
  if (sharePoint) {
    return true;
  }
  if (!sharePointScoped) {
    return false;
  }
  return api.length === 0 || api.indexOf('sharepoint') >= 0;
}

function inEnvironment(row: IRawConnection, environmentId: string): boolean {
  if (!environmentId) {
    return true;
  }
  const environment = row.properties && row.properties.environment;
  if (!environment) {
    return true;
  }
  const value = typeof environment === 'string' ? environment : `${environment.name || ''} ${environment.id || ''}`;
  const text = value.trim();
  if (!text) {
    return true;
  }
  return text === environmentId || text.indexOf(environmentId) >= 0;
}

function connectionApi(row: IRawConnection): string {
  const properties = row.properties;
  const apiId = (properties && properties.apiId) || '';
  const apiName = properties && properties.api && properties.api.name;
  return (apiName || apiId).toLowerCase();
}

function connectionIsReady(row: IRawConnection): boolean {
  const properties = row.properties;
  if (!properties || !properties.statuses || properties.statuses.length === 0) {
    return true;
  }
  return properties.statuses.some((status) => {
    const value = (status.status || '').toLowerCase();
    return value === 'connected' || value === 'ready' || value === 'ok';
  });
}

function connectionAccount(row: IRawConnection): string {
  const properties = row.properties;
  const createdBy = properties && properties.createdBy;
  return ((properties && properties.accountName) || (createdBy && createdBy.email) || '').trim().toLowerCase();
}

function connectionUser(row: IRawConnection): string {
  const properties = row.properties;
  const createdBy = properties && properties.createdBy;
  return (properties && properties.accountName)
    || (createdBy && createdBy.email)
    || (createdBy && createdBy.displayName)
    || (properties && properties.displayName)
    || row.name
    || 'SharePoint';
}

export interface IFlowRunReport {
  name: string;
  status: string;
  error: string;
  startTime: string;
}

interface IRawRunRecord {
  name?: string;
  properties?: {
    status?: string;
    startTime?: string;
    error?: { message?: string };
  };
}

interface IRawFlowAction {
  name?: string;
  properties?: {
    status?: string;
    error?: { message?: string };
  };
}

export function readLatestRun(payload: { value?: IRawRunRecord[] }): IFlowRunReport | undefined {
  const runs = (payload.value || [])
    .map(toRun)
    .filter((run): run is IFlowRunReport => run !== undefined);
  runs.sort((left, right) => right.startTime.localeCompare(left.startTime));
  return runs[0];
}

export function readFailedAction(payload: { value?: IRawFlowAction[] }): string {
  const failed = (payload.value || []).filter((action) => {
    const status = action.properties && action.properties.status;
    return status === 'Failed' || status === 'TimedOut';
  })[0];
  if (!failed) {
    return '';
  }
  const message = failed.properties && failed.properties.error && failed.properties.error.message;
  const name = (failed.name || 'Action').replace(/_/g, ' ');
  return message ? `${name}: ${message}` : name;
}

function toRun(row: IRawRunRecord): IFlowRunReport | undefined {
  if (!row.name || !row.properties || !row.properties.status) {
    return undefined;
  }
  return {
    name: row.name,
    status: row.properties.status,
    error: (row.properties.error && row.properties.error.message) || '',
    startTime: row.properties.startTime || ''
  };
}

export function readNamedFlow(
  payload: { value?: Array<{ name?: string; properties?: { displayName?: string; state?: string } }> } | undefined,
  displayName: string
): { id: string; state: FlowRunState } {
  const wanted = displayName.trim().toLowerCase();
  const rows = payload && payload.value ? payload.value : [];
  for (let index = 0; index < rows.length; index += 1) {
    const row = rows[index];
    const name = ((row.properties && row.properties.displayName) || '').trim().toLowerCase();
    if (row.name && name === wanted) {
      return { id: row.name, state: readFlowState(row) };
    }
  }
  return { id: '', state: 'Missing' };
}

export function readFlowState(payload: { name?: string; properties?: { state?: string } } | undefined): FlowRunState {
  const state = payload && payload.properties && payload.properties.state;
  if (state === 'Started' || state === 'Stopped') {
    return state;
  }
  return 'Missing';
}

export function isConsentFailure(status: number, body: string): boolean {
  const text = body.toLowerCase();
  return text.indexOf('aadsts65001') >= 0
    || text.indexOf('aadsts65002') >= 0
    || text.indexOf('has not consented') >= 0
    || text.indexOf('admin consent') >= 0
    || text.indexOf('requires consent') >= 0
    || (status === 0 && text.indexOf('consent') >= 0);
}
