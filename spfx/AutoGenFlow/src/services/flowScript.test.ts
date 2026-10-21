import { flowEditUrl } from './flowModel';
import { buildFlowPowerShell, flowApiHost, isFlowUnauthorized } from './flowScript';

describe('flow powershell fallback', () => {
  const script = buildFlowPowerShell({
    environmentId: 'Default-de2c0330-f2e8-4149-a927-0c5a8592f0da',
    environmentName: "Mohamed's default",
    apiHost: 'https://unitedarabemirates.api.flow.microsoft.com',
    connectionName: '',
    target: {
      webUrl: 'https://contoso.sharepoint.com/sites/Portal',
      siteUrl: 'https://contoso.sharepoint.com/sites/Portal',
      listGuid: 'aabbccdd-1111-2222-3333-444455556666',
      listTitle: 'tstautogenflow',
      configListTitle: 'AutoGenFlowConfiguration'
    }
  });

  it('builds a sign-in script for the selected environment and opens Edit flow', () => {
    expect(isFlowUnauthorized('api.flow.microsoft.com returned HTTP 401.')).toBe(true);
    expect(flowApiHost('https://evil.example')).toBe('https://api.flow.microsoft.com');
    expect(script).toContain("$EnvironmentId = 'Default-de2c0330-f2e8-4149-a927-0c5a8592f0da'");
    expect(script).toContain("$EnvironmentName = 'Mohamed''s default'");
    expect(script).toContain('https://unitedarabemirates.api.flow.microsoft.com');
    expect(script).toContain('Add-PowerAppsAccount');
    expect(script).toContain('Start-Process $edit');
    expect(script).not.toContain('Bearer ey');
    expect(flowEditUrl('Default-de2c0330-f2e8-4149-a927-0c5a8592f0da', 'flow-1')).toBe(
      'https://make.powerautomate.com/environments/Default-de2c0330-f2e8-4149-a927-0c5a8592f0da/flows/flow-1'
    );
    const encoded = /FromBase64String\('([^']+)'\)/.exec(script);
    expect(encoded).toBeTruthy();
    const binary = atob((encoded as RegExpExecArray)[1]);
    let json = '';
    for (let index = 0; index < binary.length; index += 1) {
      json += String.fromCharCode(binary.charCodeAt(index));
    }
    expect(json).toContain('tstautogenflow');
    expect(json).toContain('When_an_item_is_created');
    expect(json).toContain('GetOnNewItems');
    expect(json).toContain('"state":"Stopped"');
  });
});
