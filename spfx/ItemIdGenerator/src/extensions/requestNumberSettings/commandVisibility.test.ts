import { listServerRelativeUrlFromPath, shouldShowAutogenCommand } from './commandVisibility';

describe('autogen command visibility', () => {
  it('stays available before the list context is ready', () => {
    expect(shouldShowAutogenCommand(undefined, undefined, 'AutoGenFeatureConfiguration')).toBe(true);
  });

  it('hides the command on the hidden configuration list', () => {
    expect(shouldShowAutogenCommand(
      'AutoGenFeatureConfiguration',
      '/sites/ops/Lists/AutoGenFeatureConfiguration',
      'AutoGenFeatureConfiguration'
    )).toBe(false);
  });

  it('shows the command on a normal list', () => {
    expect(shouldShowAutogenCommand(
      'Requests',
      '/sites/ops/Lists/Requests',
      'AutoGenFeatureConfiguration'
    )).toBe(true);
  });
});

describe('list page path', () => {
  it('reads a list from the modern items page', () => {
    expect(listServerRelativeUrlFromPath('/sites/ops/Lists/Requests/AllItems.aspx')).toBe('/sites/ops/Lists/Requests');
  });

  it('reads a library from the forms page', () => {
    expect(listServerRelativeUrlFromPath('/sites/ops/Shared%20Documents/Forms/AllItems.aspx'))
      .toBe('/sites/ops/Shared Documents');
  });

  it('ignores site pages', () => {
    expect(listServerRelativeUrlFromPath('/sites/ops/SitePages/Home.aspx')).toBeUndefined();
  });
});
