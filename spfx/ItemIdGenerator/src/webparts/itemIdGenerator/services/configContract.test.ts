import {
  buildCreateBody,
  buildGetListUrl,
  buildListAbsoluteUrl,
  buildListItemsUrl,
  buildUpdateBody,
  deriveWebFromListUrl,
  draftForCurrentList,
  emptyDraft,
  findConfigForList,
  hasPermissionFlag,
  isConfigList,
  isExistingColumnError,
  mapConfigRow,
  PERMISSION_EDIT_LIST_ITEMS,
  validateDraft
} from './configContract';
import { buildCreateFieldBody, CONFIG_FIELD_DEFINITIONS, FIELD_CREATION_OPTIONS } from './requestNumberFields';

describe('AutoGenFeatureConfiguration contract', () => {
  const draft = {
    ...emptyDraft(),
    title: 'Operations requests',
    targetListUrl: 'https://contoso.sharepoint.com/sites/ops/Lists/Requests',
    targetListGuid: '{AAAAAAAA-BBBB-4CCC-8DDD-EEEEEEEEEEEE}',
    numberColumnInternalName: 'RequestNumber',
    formula: 'REQ-{yyyy}{MM}-{seq}',
    resetPeriod: 'Month' as const,
    isActive: true,
    padLength: 4
  };

  it('starts a new row at count 0 and does not send the counter fields', () => {
    const body = buildCreateBody(draft);
    expect(body.CurrentCount).toBe(0);
    expect(body.TargetListGuid).toBe('aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee');
    expect(body.PadLength).toBe(4);
    expect(Object.prototype.hasOwnProperty.call(body, 'LastResetDate')).toBe(false);
  });

  it('does not send the counter when an admin edits a row', () => {
    const body = buildUpdateBody({ ...draft, id: 4, currentCount: 12, lastResetDate: '2026-10-01T00:00:00Z' });
    expect(body.Title).toBe('Operations requests');
    expect(body.IsActive).toBe(true);
    expect(Object.prototype.hasOwnProperty.call(body, 'CurrentCount')).toBe(false);
    expect(Object.prototype.hasOwnProperty.call(body, 'LastResetDate')).toBe(false);
  });

  it('rejects a formula that cannot produce distinct codes and a missing GUID', () => {
    expect(validateDraft({ ...draft, formula: 'REQ-{yyyy}' })).toMatch(/\{counter\}/);
    expect(validateDraft({ ...draft, formula: 'REQ-{Title}-{counter}' })).toBeUndefined();
    expect(validateDraft({ ...draft, targetListGuid: '' })).toMatch(/Resolve list/);
    expect(validateDraft({ ...draft, padLength: 13 })).toMatch(/Pad length/);
    expect(validateDraft(draft)).toBeUndefined();
  });

  it('reads the config list with the schema field names', () => {
    const url = buildListItemsUrl('https://contoso.sharepoint.com/sites/config/', "Team's Config");
    expect(url).toContain("/lists/getbytitle('Team''s Config')/items");
    expect(url).toContain('NumberColumnInternalName');
    expect(url.indexOf('WebhookSubscriptionId')).toBe(-1);
    expect(mapConfigRow({
      Id: 4,
      Title: 'Operations requests',
      CurrentCount: 12,
      ResetPeriod: 'Month',
      IsActive: 'Yes',
      PadLength: 4,
      'odata.etag': '"3"'
    }).currentCount).toBe(12);
    expect(mapConfigRow({ Id: 4, IsActive: 'Yes' }).isActive).toBe(true);
    expect(mapConfigRow({ Id: 5 }).isActive).toBe(false);
    expect(mapConfigRow({
      Id: 6,
      Modified: '2026-10-09T11:00:00Z',
      Editor: { Title: 'Ada' }
    }).modifiedBy).toBe('Ada');
  });

  it('derives the web and builds a GetList URL', () => {
    expect(deriveWebFromListUrl('https://contoso.sharepoint.com/sites/ops/Lists/Requests/AllItems.aspx')).toEqual({
      webAbsoluteUrl: 'https://contoso.sharepoint.com/sites/ops',
      serverRelativeUrl: '/sites/ops/Lists/Requests'
    });
    const getList = buildGetListUrl("https://contoso.sharepoint.com/sites/ops/Lists/Team's Requests");
    expect(getList).toContain('https://contoso.sharepoint.com/sites/ops/_api/web/GetList(@listUrl)?@listUrl=');
    expect(decodeURIComponent(getList)).toContain("/sites/ops/Lists/Team''s Requests");
  });

  it('treats a duplicate column response as the column already being there', () => {
    expect(isExistingColumnError('A column with this name already exists. Please choose another name.')).toBe(true);
    expect(isExistingColumnError('The field name already exists.')).toBe(true);
    expect(isExistingColumnError('SharePoint returned HTTP 403.')).toBe(false);
  });

  it('describes every AutoGenFeatureConfiguration column in field XML', () => {
    const names = CONFIG_FIELD_DEFINITIONS.map((field) => field.internalName);
    expect(names).toEqual([
      'TargetListUrl',
      'TargetListGuid',
      'NumberColumnInternalName',
      'Formula',
      'CurrentCount',
      'ResetPeriod',
      'LastResetDate',
      'IsActive',
      'PadLength'
    ]);
    CONFIG_FIELD_DEFINITIONS.forEach((field) => {
      expect(field.schemaXml).toContain(`Name="${field.internalName}"`);
      const body = buildCreateFieldBody(field.schemaXml);
      expect(body.parameters.Options).toBe(FIELD_CREATION_OPTIONS);
      expect(body.parameters.SchemaXml).toBe(field.schemaXml);
      expect(Object.prototype.hasOwnProperty.call(body.parameters, '__metadata')).toBe(false);
    });
    expect(FIELD_CREATION_OPTIONS).toBe(25);
    const reset = CONFIG_FIELD_DEFINITIONS.filter((field) => field.internalName === 'ResetPeriod')[0];
    expect(reset.schemaXml).toContain('<Default>None</Default>');
    expect(emptyDraft().resetPeriod).toBe('None');
  });

  it('finds the row for the current list and hides the command on the config list', () => {
    const current = draftForCurrentList(
      'Requests',
      'https://contoso.sharepoint.com/sites/ops/Lists/Requests',
      '{AAAAAAAA-BBBB-4CCC-8DDD-EEEEEEEEEEEE}'
    );
    expect(current.targetListGuid).toBe('aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee');
    expect(findConfigForList([current, { ...emptyDraft(), targetListGuid: 'bbbbbbbb-bbbb-4ccc-8ddd-eeeeeeeeeeee' }], current.targetListGuid)?.title).toBe('Requests');
    expect(buildListAbsoluteUrl('https://contoso.sharepoint.com/sites/ops', '/sites/ops/Lists/Requests')).toBe(
      'https://contoso.sharepoint.com/sites/ops/Lists/Requests'
    );
    expect(isConfigList('AutoGenFeatureConfiguration', '/sites/ops/Lists/AutoGenFeatureConfiguration', 'AutoGenFeatureConfiguration')).toBe(true);
    expect(isConfigList('RequestNumberConfig', '/sites/ops/Lists/RequestNumberConfig', 'AutoGenFeatureConfiguration')).toBe(true);
    expect(isConfigList('Requests', '/sites/ops/Lists/Requests', 'AutoGenFeatureConfiguration')).toBe(false);
    expect(hasPermissionFlag('4', PERMISSION_EDIT_LIST_ITEMS)).toBe(true);
    expect(hasPermissionFlag('1', PERMISSION_EDIT_LIST_ITEMS)).toBe(false);
  });
});
