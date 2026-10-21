import { DEFAULT_LOG_LIST_TITLE, isAutoGenSystemList } from './contract';
import { createFlowFailureNote, createFlowLogDetail } from './flowModel';

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

  it('records the environment id when Power Automate returns no flow', () => {
    const note = createFlowFailureNote(
      { id: 'Default-de2c0330-f2e8-4149-a927-0c5a8592f0da', displayName: 'Mohamed (default)' },
      'api.flow.microsoft.com returned HTTP 401.'
    );
    expect(note).toContain('Environment id Default-de2c0330-f2e8-4149-a927-0c5a8592f0da');
    expect(note).toContain('HTTP 401');
    expect(note).toContain('https://make.powerautomate.com/environments/Default-de2c0330-f2e8-4149-a927-0c5a8592f0da/home');
    const detail = createFlowLogDetail({
      name: 'AutoGen Flow · tstautogenflow',
      environmentId: 'Default-de2c0330-f2e8-4149-a927-0c5a8592f0da',
      flowId: '',
      flowState: '',
      note
    });
    expect(detail).toContain('Id none');
    expect(detail).toContain('State none');
    expect(detail).toContain('Environment Default-de2c0330-f2e8-4149-a927-0c5a8592f0da');
  });
});