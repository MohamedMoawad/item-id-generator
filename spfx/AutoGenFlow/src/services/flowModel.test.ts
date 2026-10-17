import { chooseEnvironment, isConsentFailure, readConnections, readFlowState } from './flowModel';

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
      displayName: 'Default'
    });
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
