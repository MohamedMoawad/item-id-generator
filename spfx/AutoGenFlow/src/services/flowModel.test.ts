import { chooseEnvironment, connectionListUrls, isConsentFailure, readConnections, readEnvironments, readFlowState } from './flowModel';

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
      powerAppsEndpoint: 'https://api.powerapps.com'
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
    expect(urls[0]).toContain('/providers/Microsoft.PowerApps/apis/shared_sharepointonline/connections');
    expect(urls[0]).toContain('Default-tenant');
    expect(urls[urls.length - 1]).toContain('/providers/Microsoft.ProcessSimple/environments/Default-tenant/connections');
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
    })).toEqual([{ name: 'shared-sharepointonl-1', displayName: 'SharePoint' }]);
  });

  it('reads the flow state and treats a missing approval as consent', () => {
    expect(readFlowState({ properties: { state: 'Started' } })).toBe('Started');
    expect(readFlowState({})).toBe('Missing');
    expect(isConsentFailure(401, '')).toBe(true);
    expect(isConsentFailure(400, 'The column is missing')).toBe(false);
  });
});
