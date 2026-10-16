import {
  buildCreateBody,
  buildUpdateBody,
  DEFAULT_CONFIG_LIST_TITLE,
  draftForCurrentList,
  findConfigForList,
  isFlowConfigList,
  validateDraft
} from './contract';
import { CONFIG_FIELD_DEFINITIONS } from './fields';
import { prepareRule } from './formula';

describe('contract', () => {
  const draft = {
    ...draftForCurrentList('Requests', 'https://contoso.sharepoint.com/sites/ops/Lists/Requests', '{AABBCCDD-1111-2222-3333-444455556666}'),
    numberColumnInternalName: 'RequestNumber'
  };

  it('uses a master list that is separate from AutoGen Feature', () => {
    expect(DEFAULT_CONFIG_LIST_TITLE).toBe('AutoGenFlowConfiguration');
    expect(isFlowConfigList('AutoGenFlowConfiguration', '/sites/ops/Lists/AutoGenFlowConfiguration', DEFAULT_CONFIG_LIST_TITLE)).toBe(true);
    expect(isFlowConfigList('AutoGenFeatureConfiguration', '/sites/ops/Lists/AutoGenFeatureConfiguration', DEFAULT_CONFIG_LIST_TITLE)).toBe(false);
    const names = CONFIG_FIELD_DEFINITIONS.map((field) => field.internalName);
    expect(names).toEqual(expect.arrayContaining([
      'TargetListGuid',
      'FormulaColumns',
      'CounterToken',
      'CounterWidth',
      'CurrentCount'
    ]));
  });

  it('stores the prepared formula and leaves the live counter alone on a later save', () => {
    const prepared = prepareRule(draft.formula, draft.padLength);
    const row = { ...draft, ...prepared, currentCount: 12 };
    expect(validateDraft(row)).toBeUndefined();
    expect(buildCreateBody({ ...draft, ...prepared }).CurrentCount).toBe(0);
    expect(buildCreateBody(row).CurrentCount).toBe(12);
    expect(buildUpdateBody(row, false).CurrentCount).toBeUndefined();
    expect(buildUpdateBody(row, true).CurrentCount).toBe(12);
    expect(buildUpdateBody(row, true).CounterToken).toBe('{counter:4}');
    expect(findConfigForList([row], 'aabbccdd-1111-2222-3333-444455556666')).toBe(row);
  });
});
