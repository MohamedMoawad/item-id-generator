import { DEFAULT_LOG_LIST_TITLE, isAutoGenSystemList } from './contract';

describe('flow log list', () => {
  it('keeps the log list out of numbering', () => {
    expect(DEFAULT_LOG_LIST_TITLE).toBe('AutoGenFlowLog');
    expect(isAutoGenSystemList(
      'AutoGenFlowLog',
      '/sites/ops/Lists/AutoGenFlowLog',
      'AutoGenFlowConfiguration'
    )).toBe(true);
    expect(isAutoGenSystemList(
      'AutoGenFlowConfiguration',
      '/sites/ops/Lists/AutoGenFlowConfiguration',
      'AutoGenFlowConfiguration'
    )).toBe(true);
    expect(isAutoGenSystemList(
      'Requests',
      '/sites/ops/Lists/Requests',
      'AutoGenFlowConfiguration'
    )).toBe(false);
  });
});