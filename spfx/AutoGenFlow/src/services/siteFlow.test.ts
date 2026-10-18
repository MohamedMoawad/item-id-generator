import { buildCreateFlowPayload, buildShellFlowPayload, buildSiteFlowDefinition, flowDisplayName, IListFlowTarget } from './siteFlow';

const target: IListFlowTarget = {
  webUrl: 'https://contoso.sharepoint.com/sites/PortalTeam/subsite',
  siteUrl: 'https://contoso.sharepoint.com/sites/PortalTeam/',
  listGuid: '{AABBCCDD-1111-2222-3333-444455556666}',
  listTitle: 'Requests',
  configListTitle: 'AutoGenFlowConfiguration'
};

describe('list flow', () => {
  const text = JSON.stringify(buildSiteFlowDefinition(target));

  it('creates one flow for the list and reads the site collection rule', () => {
    expect(text).toContain('GetOnUpdatedItems');
    expect(text).toContain('"dataset":"https://contoso.sharepoint.com/sites/PortalTeam/subsite"');
    expect(text).toContain('"dataset":"https://contoso.sharepoint.com/sites/PortalTeam"');
    expect(text).toContain('"table":"aabbccdd-1111-2222-3333-444455556666"');
    expect(text).toContain("TargetListGuid%20eq%20'aabbccdd-1111-2222-3333-444455556666'");
    expect(text).toContain("body('Get_rule')?['d']?['results']");
    expect(text).toContain('Column_name');
    expect(text).toContain('No AutoGen rule was found for this list.');
    expect(text).not.toContain('+eq+');
    expect(text).toContain('AutoGenFlowConfiguration');
    expect(text).toContain('ItemQuery');
    expect(text).toContain('"runs":1');
    expect(text).toContain('"repetitions":1');
    expect(text).not.toContain('"frequency":"Minute"');
    expect(flowDisplayName('Requests')).toBe('AutoGen · Requests');
  });

  it('binds the SharePoint connection the person chose', () => {
    const payload = buildCreateFlowPayload(target, 'shared-sharepointonl-1');
    const connection = payload.properties.connectionReferences.shared_sharepointonline;
    expect(connection && connection.connectionName).toBe('shared-sharepointonl-1');
    expect(payload.properties.displayName).toBe('AutoGen · Requests');
    expect(payload.properties.state).toBe('Stopped');
    const unbound = buildCreateFlowPayload(target, '');
    expect(unbound.properties.connectionReferences.shared_sharepointonline).toBeUndefined();
    const shell = JSON.stringify(buildShellFlowPayload('Requests'));
    expect(shell).toContain('AutoGen · Requests');
    expect(shell).toContain('"state":"Stopped"');
    expect(shell).not.toContain('shared_sharepointonline');
    expect(shell).not.toContain('GetOnUpdatedItems');
  });
});
