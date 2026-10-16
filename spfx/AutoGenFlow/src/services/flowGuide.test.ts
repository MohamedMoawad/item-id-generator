import {
  blankGuardExpression,
  CODE_START_EXPRESSION,
  configFilter,
  FLOW_CONCURRENCY,
  flowExpressions,
  PADDED_EXPRESSION,
  ROLLED_EXPRESSION,
  triggerCondition
} from './flowGuide';

describe('flow guide', () => {
  it('binds the flow to one list and keeps runs in a single line', () => {
    expect(FLOW_CONCURRENCY).toBe(1);
    expect(configFilter('{AABBCCDD-1111-2222-3333-444455556666}')).toBe(
      "TargetListGuid eq 'aabbccdd-1111-2222-3333-444455556666' and IsActive eq 1"
    );
    expect(triggerCondition('RequestNumber')).toBe("@empty(triggerOutputs()?['body/RequestNumber'])");
    expect(blankGuardExpression('RequestNumber')).toContain("outputs('Get_item')?['body/RequestNumber']");
  });

  it('guards date formatting and pads from twelve zeros', () => {
    expect(ROLLED_EXPRESSION.startsWith('if(')).toBe(true);
    expect(ROLLED_EXPRESSION).toContain("equals(outputs('ResetPeriod'), 'None')");
    expect(PADDED_EXPRESSION).toContain("'000000000000'");
    expect(CODE_START_EXPRESSION).toContain("outputs('CounterToken')");
    expect(CODE_START_EXPRESSION).toContain("'{MM}'");
    const names = flowExpressions('RequestNumber').map((item) => item.action);
    expect(names).toContain('Next');
    expect(names).toContain('Padded');
    expect(names).toContain('CodeStart');
  });
});
