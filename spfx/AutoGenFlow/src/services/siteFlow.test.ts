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

  it('creates an automated flow for this list when an item is created', () => {
    expect(text).toContain('GetOnNewItems');
    expect(text).toContain('When_an_item_is_created');
    expect(text).toContain('"dataset":"https://contoso.sharepoint.com/sites/PortalTeam/subsite"');
    expect(text).toContain('"table":"aabbccdd-1111-2222-3333-444455556666"');
    expect(text).toContain("@triggerBody()?['value']");
    expect(text).toContain('AutoGenFlowConfiguration');
    expect(text).toContain('TargetListGuid%20eq%20');
    expect(text).toContain("body('Get_rule')?['d']?['results']");
    expect(text).toContain('@triggerBody()');
    expect(text).toContain("add(variables('Previous'), 1)");
    expect(text).toContain('Column_name');
    expect(text).toContain('"runs":1');
    expect(text).not.toContain('GetOnUpdatedItems');
    expect(text).not.toContain('For_each_rule');
    expect(flowDisplayName(target.listTitle)).toBe('AutoGen Flow · Requests');
  });

  it('binds the SharePoint connection the person chose', () => {
    const payload = buildCreateFlowPayload(target, 'shared-sharepointonl-1');
    const connection = payload.properties.connectionReferences.shared_sharepointonline;
    expect(connection && connection.connectionName).toBe('shared-sharepointonl-1');
    expect(payload.properties.displayName).toBe('AutoGen Flow · Requests');
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
    expect(row.name).toBe('AutoGen Flow · Requests');
    expect(JSON.stringify(clientdata.properties.definition)).toContain('GetOnNewItems');
    expect(JSON.stringify(clientdata.properties.definition)).not.toContain('GetOnUpdatedItems');
    expect(clientdata.properties.connectionReferences.shared_sharepointonline.connection.name).toBe('shared-sharepointonl-1');
    const portal = JSON.parse(buildPortalTemplate(target, 'shared-sharepointonl-1')) as { displayName: string; properties: { definition: object; connectionReferences: { shared_sharepointonline: { connectionName: string } } } };
    expect(portal.displayName).toBe('AutoGen Flow · Requests');
    expect(JSON.stringify(portal.properties.definition)).toContain('GetOnNewItems');
    expect(JSON.stringify(portal.properties.definition)).not.toContain('GetOnUpdatedItems');
    expect(portal.properties.connectionReferences.shared_sharepointonline.connectionName).toBe('shared-sharepointonl-1');
  });
});
