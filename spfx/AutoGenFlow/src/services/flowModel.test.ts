import { chooseConnection, chooseEnvironment, connectionListUrls, connectionPortalUrl, flowCreateUrls, flowPutUrl, flowGeoCandidates, flowHomeUrl, flowListUrl, flowPortalUrl, isConsentFailure, publicFlowRedirect, readConnectionNamesFromFlows, readConnections, readCreatedConnection, readEnvironments, readFailedAction, readFlowState, readLatestRun, readNamedFlow, sharePointConnectionBody, sharePointConnectionCreateUrls, sharePointConnectionPutUrl } from './flowModel';

describe('flow model', () => {
  it('uses the default environment endpoint and SharePoint connections', () => {
    expect(chooseEnvironment({
      value: [
        { name: 'other', properties: { isDefault: false, displayName: 'Other' } },
        {
          name: 'Default-tenant',
          properties: {
            isDefault: true,
            displayName: 'Default',
            runtimeEndpoints: { 'microsoft.Flow': 'https://unitedstates.api.flow.microsoft.com/' }
          }
        }
      ]
    })).toEqual({
      id: 'Default-tenant',
      endpoint: 'https://unitedstates.api.flow.microsoft.com',
      displayName: 'Default',
      isDefault: true,
      powerAppsEndpoint: 'https://api.powerapps.com',
      location: 'unitedstates',
      instanceUrl: ''
    });
    const environment = chooseEnvironment({
      value: [{
        name: 'Default-tenant',
        properties: {
          isDefault: true,
          displayName: 'Default',
          runtimeEndpoints: { 'microsoft.Flow': 'https://unitedstates.api.flow.microsoft.com/' }
        }
      }]
    });
    const urls = environment ? connectionListUrls(environment) : [];
    expect(urls[0]).toContain('https://api.powerapps.com/providers/Microsoft.PowerApps/connections?');
    expect(urls[0]).toContain('Default-tenant');
    expect(urls[0]).not.toContain('/apis/shared_sharepointonline');
    expect(urls[urls.length - 1]).toContain('/providers/Microsoft.ProcessSimple/environments/Default-tenant/connections');
    const regional = connectionListUrls({
      id: 'Default-tenant',
      endpoint: 'https://unitedstates.api.flow.microsoft.com',
      powerAppsEndpoint: 'https://unitedstates.api.powerapps.com',
      displayName: 'Default',
      isDefault: true,
      location: 'unitedstates',
      instanceUrl: ''
    });
    expect(regional.some((url) => url.indexOf('https://api.powerapps.com/providers/Microsoft.PowerApps/connections?') === 0)).toBe(true);
    expect(chooseEnvironment({
      value: [{
        name: 'Default-de2c0330-f2e8-4149-a927-0c5a8592f0da',
        properties: {
          isDefault: true,
          displayName: 'Personal Productivity(Default)',
          runtimeEndpoints: {
            'microsoft.Flow': 'https://10.0.5.41:21497',
            'microsoft.PowerApps': 'https://10.0.5.41:21498'
          }
        }
      }]
    })).toEqual({
      id: 'Default-de2c0330-f2e8-4149-a927-0c5a8592f0da',
      endpoint: 'https://api.flow.microsoft.com',
      powerAppsEndpoint: 'https://api.powerapps.com',
      displayName: 'Personal Productivity(Default)',
      isDefault: true,
      location: '',
      instanceUrl: ''
    });
    expect(readEnvironments({
      value: [
        { name: 'prd', properties: { isDefault: false, displayName: 'SharePoint (PRD)' } },
        {
          name: 'Default-tenant',
          properties: {
            isDefault: true,
            displayName: 'Default',
            runtimeEndpoints: { 'microsoft.Flow': 'https://unitedstates.api.flow.microsoft.com/' }
          }
        }
      ]
    }).map((item) => item.displayName)).toEqual(['Default', 'SharePoint (PRD)']);
    expect(readConnections({
      value: [
        { name: 'shared-sharepointonl-1', properties: { displayName: 'SharePoint', apiId: '/providers/Microsoft.PowerApps/apis/shared_sharepointonline', statuses: [{ status: 'Connected' }] } },
        { name: 'sql-1', properties: { displayName: 'SQL', apiId: '/providers/Microsoft.PowerApps/apis/shared_sql' } }
      ]
    })).toEqual([{ name: 'shared-sharepointonl-1', displayName: 'SharePoint', accountName: '' }]);
    expect(readConnections({
      value: [{
        name: 'shared-sharepointonl-2',
        properties: {
          displayName: 'SharePoint',
          accountName: 'svc@contoso.com',
          apiId: '/providers/Microsoft.PowerApps/apis/shared_sharepointonline',
          statuses: [{ status: 'Error' }]
        }
      }]
    })).toEqual([{ name: 'shared-sharepointonl-2', displayName: 'svc@contoso.com', accountName: 'svc@contoso.com' }]);
    const connections = readConnections({
      value: [
        { name: 'shared-sharepointonl-1', properties: { displayName: 'SharePoint', accountName: 'other@contoso.com', apiId: '/providers/Microsoft.PowerApps/apis/shared_sharepointonline', statuses: [{ status: 'Connected' }] } },
        { name: 'shared-sharepointonl-2', properties: { displayName: 'SharePoint', accountName: 'me@contoso.com', apiId: '/providers/Microsoft.PowerApps/apis/shared_sharepointonline', statuses: [{ status: 'Connected' }] } }
      ]
    });
    const mine = chooseConnection(connections, 'i:0#.f|membership|me@contoso.com');
    const first = chooseConnection(connections, '');
    expect(mine && mine.name).toBe('shared-sharepointonl-2');
    expect(first && first.name).toBe(connections[0].name);
    expect(flowPortalUrl('Default-tenant', 'flow-1')).toBe('https://make.powerautomate.com/environments/Default-tenant/flows/flow-1/details');
    const createUrls = flowCreateUrls({
      id: '2ab1148e-14ab-e29a-acaf-fc1efecf9050',
      location: 'unitedarabemirates',
      endpoint: 'https://api.flow.microsoft.com'
    });
    expect(createUrls[0]).toBe('https://unitedarabemirates.api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/2ab1148e-14ab-e29a-acaf-fc1efecf9050/flows?api-version=2016-11-01');
    expect(createUrls[1]).toContain('https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/2ab1148e-14ab-e29a-acaf-fc1efecf9050/flows?');
    expect(createUrls.join(' ')).not.toContain('europe.api');
    expect(createUrls.join(' ')).not.toContain('scopes/admin');
    expect(createUrls.join(' ')).not.toContain('10.');
    const fromGlobal = 'https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/env/flows?api-version=2016-11-01';
    expect(publicFlowRedirect(fromGlobal, 'https://unitedarabemirates.api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/env/flows?api-version=2016-11-01')).toBe('');
    expect(publicFlowRedirect(fromGlobal, 'https://10.0.5.41:21497/providers/Microsoft.ProcessSimple/environments/env/flows?api-version=2016-11-01')).toBe('');
    expect(publicFlowRedirect('https://unitedarabemirates.api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/env/flows?api-version=2016-11-01', '/providers/Microsoft.ProcessSimple/environments/env/flows?api-version=2016-11-01'))
      .toBe('https://unitedarabemirates.api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/env/flows?api-version=2016-11-01');
    const located = chooseEnvironment({
      value: [{
        name: 'Default-de2c0330-f2e8-4149-a927-0c5a8592f0da',
        location: 'unitedarabemirates',
        properties: {
          isDefault: true,
          displayName: 'Personal Productivity(Default)',
          runtimeEndpoints: { 'microsoft.Flow': 'https://10.0.5.41:21497' }
        }
      }]
    });
    const locatedUrls = located ? flowCreateUrls(located) : [];
    expect(locatedUrls[0]).toContain('https://unitedarabemirates.api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/Default-de2c0330-f2e8-4149-a927-0c5a8592f0da/flows?');
    expect(locatedUrls.join(' ')).not.toContain('scopes/admin');
    expect(locatedUrls.join(' ')).not.toContain('10.');
    const qatar = chooseEnvironment({
      value: [{
        name: '2ab1148e-14ab-e29a-acaf-fc1efecf9050',
        properties: {
          isDefault: false,
          displayName: 'Qatargas',
          azureRegionHint: 'uaenorth',
          runtimeEndpoints: { 'microsoft.Flow': 'https://10.0.5.41:21497' },
          linkedEnvironmentMetadata: { instanceUrl: 'https://org.crm15.dynamics.com' }
        }
      }]
    });
    expect(qatar && qatar.location).toBe('unitedarabemirates');
    expect(qatar && qatar.instanceUrl).toBe('https://org.crm15.dynamics.com');
    expect(qatar && flowCreateUrls(qatar).join(' ')).not.toContain('europe.api');
    const labeledEurope = chooseEnvironment({
      value: [{
        name: '2ab1148e-14ab-e29a-acaf-fc1efecf9050',
        location: 'europe',
        properties: {
          displayName: 'Qatar',
          azureRegionHint: 'qatarcentral',
          runtimeEndpoints: { 'microsoft.Flow': 'https://europe.api.flow.microsoft.com' },
          linkedEnvironmentMetadata: { instanceUrl: 'https://org.crm15.dynamics.com' }
        }
      }]
    });
    expect(labeledEurope && labeledEurope.location).toBe('unitedarabemirates');
    expect(labeledEurope && flowCreateUrls(labeledEurope).join(' ')).not.toContain('europe.api');
    const europeLabel = chooseEnvironment({
      value: [{
        name: 'Default-de2c0330-f2e8-4149-a927-0c5a8592f0da',
        location: 'europe',
        properties: {
          displayName: 'Personal Productivity(Default)',
          azureRegionHint: 'westeurope',
          runtimeEndpoints: { 'microsoft.Flow': 'https://europe.api.flow.microsoft.com' }
        }
      }]
    });
    expect(europeLabel && europeLabel.location).toBe('europe');
    expect(europeLabel && flowCreateUrls(europeLabel)[0]).toContain('https://emea.api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/Default-de2c0330-f2e8-4149-a927-0c5a8592f0da/flows?');
    expect(europeLabel && flowCreateUrls(europeLabel)[1]).toContain('https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/Default-de2c0330-f2e8-4149-a927-0c5a8592f0da/flows?');
    expect(europeLabel && flowCreateUrls(europeLabel).join(' ')).not.toContain('europe.api');
    expect(europeLabel && flowCreateUrls(europeLabel).join(' ')).not.toContain('unitedarabemirates');
    expect(flowPutUrl('https://emea.api.flow.microsoft.com', 'Default-de2c0330-f2e8-4149-a927-0c5a8592f0da', '8f3e2a10-6c4b-4d11-9a77-1b2c3d4e5f60'))
      .toBe('https://emea.api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/Default-de2c0330-f2e8-4149-a927-0c5a8592f0da/flows/8f3e2a10-6c4b-4d11-9a77-1b2c3d4e5f60?api-version=2016-11-01');
    expect(qatar && flowGeoCandidates(qatar)[0]).toBe('https://unitedarabemirates.api.flow.microsoft.com');
    expect(qatar && flowGeoCandidates(qatar).join(' ')).not.toContain('10.');
    expect(qatar && flowGeoCandidates(qatar).join(' ')).not.toContain('scopes/admin');
    expect(flowHomeUrl('Default-de2c0330-f2e8-4149-a927-0c5a8592f0da')).toBe('https://make.powerautomate.com/environments/Default-de2c0330-f2e8-4149-a927-0c5a8592f0da/home');
    expect(flowListUrl('Default-de2c0330-f2e8-4149-a927-0c5a8592f0da')).toBe('https://make.powerautomate.com/environments/Default-de2c0330-f2e8-4149-a927-0c5a8592f0da/flows');
    expect(readConnectionNamesFromFlows({
      value: [{
        properties: {
          connectionReferences: {
            shared_sharepointonline: {
              connectionName: 'shared-sharepointonl-7',
              id: '/providers/Microsoft.PowerApps/apis/shared_sharepointonline'
            },
            shared_office365: { connectionName: 'shared-office365-1', id: '/providers/Microsoft.PowerApps/apis/shared_office365' }
          }
        }
      }]
    })).toEqual([{ name: 'shared-sharepointonl-7', displayName: 'shared-sharepointonl-7', accountName: '' }]);
    expect(readConnectionNamesFromFlows({
      properties: {
        connectionReferences: {
          shared_sharepointonline: { connection: { name: 'shared-sharepointonl-8' }, api: { name: 'shared_sharepointonline' } }
        }
      }
    })[0].name).toBe('shared-sharepointonl-8');
    expect(readConnections({
      value: [{
        id: '/providers/Microsoft.PowerApps/apis/shared_sharepointonline/connections/shared-sharepointonl-9',
        properties: { displayName: 'SharePoint', environment: { id: 'Default-tenant' } }
      }]
    }, false, 'Default-tenant')[0].name).toBe('shared-sharepointonl-9');
  });

  it('creates a SharePoint connection from the flow host', () => {
    const environment = {
      id: 'Default-tenant',
      endpoint: 'https://unitedstates.api.flow.microsoft.com',
      powerAppsEndpoint: 'https://unitedstates.api.powerapps.com',
      displayName: 'Default',
      isDefault: true,
      location: 'unitedstates',
      instanceUrl: ''
    };
    const urls = sharePointConnectionCreateUrls(environment);
    expect(urls[0]).toContain('https://api.flow.microsoft.com/providers/Microsoft.PowerApps/apis/shared_sharepointonline/connections?');
    expect(urls[0]).toContain('Default-tenant');
    expect(sharePointConnectionPutUrl(environment, 'shared-sharepointonl-new')).toContain('/connections/shared-sharepointonl-new?');
    expect(sharePointConnectionBody('Default-tenant').properties.environment.id).toBe('/providers/Microsoft.PowerApps/environments/Default-tenant');
    expect(readCreatedConnection({
      name: 'shared-sharepointonl-4',
      properties: { displayName: 'SharePoint', accountName: 'me@contoso.com', statuses: [{ status: 'Connected' }] }
    })).toEqual({
      name: 'shared-sharepointonl-4',
      displayName: 'me@contoso.com',
      accountName: 'me@contoso.com',
      consentUrl: '',
      connected: true
    });
    expect(readCreatedConnection({
      name: 'shared-sharepointonl-5',
      properties: {
        displayName: 'SharePoint',
        statuses: [{ status: 'Unauthenticated' }],
        consentInfo: { redirectUrl: 'https://login.microsoftonline.com/consent' }
      }
    })).toEqual({
      name: 'shared-sharepointonl-5',
      displayName: 'SharePoint',
      accountName: '',
      consentUrl: 'https://login.microsoftonline.com/consent',
      connected: false
    });
    expect(connectionPortalUrl('Default-tenant')).toBe('https://make.powerautomate.com/environments/Default-tenant/connections');
  });

  it('reads the flow state and treats a missing approval as consent', () => {
    expect(readFlowState({ properties: { state: 'Started' } })).toBe('Started');
    expect(readFlowState({})).toBe('Missing');
    expect(readNamedFlow({
      value: [
        { name: 'other-flow', properties: { displayName: 'Other', state: 'Started' } },
        { name: 'site-flow', properties: { displayName: 'AutoGen Flow · PortalTeam', state: 'Started' } }
      ]
    }, 'AutoGen Flow · PortalTeam')).toEqual({ id: 'site-flow', state: 'Started' });
    expect(readNamedFlow({ value: [] }, 'AutoGen Flow · PortalTeam')).toEqual({ id: '', state: 'Missing' });
    expect(isConsentFailure(401, '')).toBe(false);
    expect(isConsentFailure(403, 'The user is not authorized for this environment.')).toBe(false);
    expect(isConsentFailure(0, 'AADSTS65001: The user or administrator has not consented to use the application.')).toBe(true);
    expect(isConsentFailure(400, 'The column is missing')).toBe(false);
  });

  it('reads the newest flow run and the failed action', () => {
    expect(readLatestRun({
      value: [
        { name: 'older', properties: { status: 'Failed', startTime: '2026-01-01T00:00:00Z', error: { message: 'old' } } },
        { name: 'newer', properties: { status: 'Succeeded', startTime: '2026-02-01T00:00:00Z' } }
      ]
    })).toEqual({ name: 'newer', status: 'Succeeded', error: '', startTime: '2026-02-01T00:00:00Z' });
    expect(readFailedAction({
      value: [
        { name: 'Get_rule', properties: { status: 'Succeeded' } },
        { name: 'Write_item', properties: { status: 'Failed', error: { message: 'The field was not found.' } } }
      ]
    })).toBe('Write item: The field was not found.');
  });
});
