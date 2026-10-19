import { buildCreateFlowPayload, buildDataverseWorkflowBody, buildPortalTemplate, buildShellFlowPayload, buildSiteFlowDefinition, flowDisplayName, IListFlowTarget } from './siteFlow';

const target: IListFlowTarget = {
  webUrl: 'https://contoso.sharepoint.com/sites/PortalTeam/subsite',
  siteUrl: 'https://contoso.sharepoint.com/sites/PortalTeam/',
  listGuid: '{AABBCCDD-1111-2222-3333-444455556666}',
  listTitle: 'Requests',
  configListTitle: 'AutoGenFlowConfiguration'
};

describe('site flow', () => {
  const text = JSON.stringify(buildSiteFlowDefinition(target));

  it('creates one flow for the site collection and reads every active rule', () => {
    expect(text).toContain('"frequency":"Minute"');
    expect(text).toContain('"interval":1');
    expect(text).toContain('"dataset":"https://contoso.sharepoint.com/sites/PortalTeam"');
    expect(text).toContain('AutoGenFlowConfiguration');
    expect(text).toContain('TargetListUrl');
    expect(text).toContain('For_each_rule');
    expect(text).toContain('For_each_item');
    expect(text).toContain("IsActive%20eq%201%20and%20TargetListGuid%20ne%20'site'");
    expect(text).toContain("body('Get_rules')?['d']?['results']");
    expect(text).toContain("@outputs('Web')");
    expect(text).toContain('ItemQuery');
    expect(text).toContain('Modified%20desc');
    expect(text).toContain("add(variables('Previous'), 1)");
    expect(text).toContain('Column_name');
    expect(text).toContain('"runs":1');
    expect(text).toContain('"repetitions":1');
    expect(text).not.toContain('GetOnUpdatedItems');
    expect(text).not.toContain('aabbccdd-1111-2222-3333-444455556666');
    expect(text).not.toContain('subsite');
    expect(text).not.toContain('No AutoGen rule was found for this list.');
    expect(flowDisplayName(target.siteUrl)).toBe('AutoGen Flow · PortalTeam');
  });

  it('binds the SharePoint connection the person chose', () => {
    const payload = buildCreateFlowPayload(target, 'shared-sharepointonl-1');
    const connection = payload.properties.connectionReferences.shared_sharepointonline;
    expect(connection && connection.connectionName).toBe('shared-sharepointonl-1');
    expect(payload.properties.displayName).toBe('AutoGen Flow · PortalTeam');
    expect(payload.properties.state).toBe('Stopped');
    const unbound = buildCreateFlowPayload(target, '');
    expect(unbound.properties.connectionReferences.shared_sharepointonline).toBeUndefined();
    const shell = JSON.stringify(buildShellFlowPayload('Requests'));
    expect(shell).toContain('AutoGen Flow · Requests');
    expect(shell).toContain('"state":"Stopped"');
    expect(shell).not.toContain('shared_sharepointonline');
    expect(shell).not.toContain('GetOnUpdatedItems');
    expect(shell).not.toContain('For_each_rule');
    const row = buildDataverseWorkflowBody(target, 'shared-sharepointonl-1');
    const clientdata = JSON.parse(row.clientdata) as { properties: { definition: { triggers: object }; connectionReferences: { shared_sharepointonline: { connection: { name: string } } } } };
    expect(row.category).toBe(5);
    expect(row.primaryentity).toBe('none');
    expect(row.name).toBe('AutoGen Flow · PortalTeam');
    expect(JSON.stringify(clientdata.properties.definition)).toContain('"frequency":"Minute"');
    expect(JSON.stringify(clientdata.properties.definition)).not.toContain('GetOnUpdatedItems');
    expect(clientdata.properties.connectionReferences.shared_sharepointonline.connection.name).toBe('shared-sharepointonl-1');
    const portal = JSON.parse(buildPortalTemplate(target, 'shared-sharepointonl-1')) as { displayName: string; properties: { definition: object; connectionReferences: { shared_sharepointonline: { connectionName: string } } } };
    expect(portal.displayName).toBe('AutoGen Flow · PortalTeam');
    expect(JSON.stringify(portal.properties.definition)).toContain('For_each_rule');
    expect(JSON.stringify(portal.properties.definition)).not.toContain('GetOnUpdatedItems');
    expect(portal.properties.connectionReferences.shared_sharepointonline.connectionName).toBe('shared-sharepointonl-1');
  });
});
