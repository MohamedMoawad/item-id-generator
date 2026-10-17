import { buildCreateFlowPayload, buildSiteFlowDefinition, flowDisplayName } from './siteFlow';

describe('site flow', () => {
  const definition = buildSiteFlowDefinition('https://contoso.sharepoint.com/sites/PortalTeam/', 'AutoGenFlowConfiguration');
  const text = JSON.stringify(definition);

  it('numbers every saved list from one site flow', () => {
    expect(text).toContain('https://contoso.sharepoint.com/sites/PortalTeam');
    expect(text).toContain("TargetListGuid+ne+'site'");
    expect(text).toContain('ItemQuery');
    expect(text).toContain('"frequency":"Minute"');
    expect(text).toContain('"runs":1');
    expect(text).toContain('"repetitions":1');
    expect(flowDisplayName('https://contoso.sharepoint.com/sites/PortalTeam')).toBe('AutoGen · PortalTeam');
  });

  it('binds the SharePoint connection the person chose', () => {
    const payload = buildCreateFlowPayload(
      'https://contoso.sharepoint.com/sites/PortalTeam',
      'AutoGenFlowConfiguration',
      'shared-sharepointonl-1'
    );
    expect(payload.properties.connectionReferences.shared_sharepointonline.connectionName).toBe('shared-sharepointonl-1');
    expect(payload.properties.displayName).toBe('AutoGen · PortalTeam');
  });
});
