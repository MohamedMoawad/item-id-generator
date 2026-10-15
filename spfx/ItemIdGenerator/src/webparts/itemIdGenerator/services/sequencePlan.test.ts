import { counterIsCurrent, planContinuousNumbers } from './sequencePlan';

const now = new Date('2026-10-15T12:00:00Z');

describe('continuous request numbers', () => {
  it('numbers the first item through the last, including blanks and gaps', () => {
    const plan = planContinuousNumbers({
      items: [
        { id: 30, created: '2026-10-15T09:00:00Z', current: 'REQ-05' },
        { id: 10, created: '2026-10-15T08:00:00Z', current: '' },
        { id: 20, created: '2026-10-15T08:30:00Z', current: 'REQ-01' }
      ],
      formula: 'REQ-{counter:2}',
      resetPeriod: 'None',
      padLength: 0,
      now
    });
    expect(plan.plans.map((item) => item.code)).toEqual(['REQ-01', 'REQ-02', 'REQ-03']);
    expect(plan.plans.map((item) => item.id)).toEqual([10, 20, 30]);
    expect(plan.plans.map((item) => item.changed)).toEqual([true, true, true]);
    expect(plan.currentCount).toBe(3);
  });

  it('leaves an already continuous list unchanged', () => {
    const plan = planContinuousNumbers({
      items: [
        { id: 1, created: '2026-10-15T08:00:00Z', current: 'REQ-01' },
        { id: 2, created: '2026-10-15T08:01:00Z', current: 'REQ-02' }
      ],
      formula: 'REQ-{counter:2}',
      resetPeriod: 'None',
      padLength: 0,
      now
    });
    expect(plan.plans.every((item) => !item.changed)).toBe(true);
    expect(plan.currentCount).toBe(2);
  });

  it('restarts the sequence for the current reset period', () => {
    const plan = planContinuousNumbers({
      items: [
        { id: 1, created: '2026-09-30T08:00:00Z', current: 'OLD-01' },
        { id: 2, created: '2026-10-15T08:00:00Z', current: '' },
        { id: 3, created: '2026-10-15T09:00:00Z', current: 'REQ-04' }
      ],
      formula: 'REQ-{counter:2}',
      resetPeriod: 'Month',
      padLength: 0,
      now
    });
    expect(plan.plans.map((item) => ({ id: item.id, code: item.code }))).toEqual([
      { id: 2, code: 'REQ-01' },
      { id: 3, code: 'REQ-02' }
    ]);
    expect(plan.currentCount).toBe(2);
  });

  it('uses each item column when the formula asks for it', () => {
    const plan = planContinuousNumbers({
      items: [
        { id: 4, created: '2026-10-15T08:00:00Z', current: '', columns: { Title: 'Desk' } },
        { id: 5, created: '2026-10-15T08:05:00Z', current: '', columns: { Title: 'Chair' } }
      ],
      formula: '{Title}-{counter:2}',
      resetPeriod: 'None',
      padLength: 0,
      now
    });
    expect(plan.plans.map((item) => item.code)).toEqual(['Desk-01', 'Chair-02']);
  });

  it('rewrites 00011, 00022, and 00033 as 00001, 00002, and 00003', () => {
    const plan = planContinuousNumbers({
      items: [
        { id: 1, created: '2026-10-15T08:00:00Z', current: '00011' },
        { id: 2, created: '2026-10-15T08:01:00Z', current: '00022' },
        { id: 3, created: '2026-10-15T08:02:00Z', current: '00033' }
      ],
      formula: '{counter:5}',
      resetPeriod: 'None',
      padLength: 0,
      now
    });
    expect(plan.plans.map((item) => item.code)).toEqual(['00001', '00002', '00003']);
    expect(plan.plans.every((item) => item.changed)).toBe(true);
    expect(plan.currentCount).toBe(3);
  });

  it('does not append the item id or a second counter token', () => {
    const plan = planContinuousNumbers({
      items: [
        { id: 1, created: '2026-10-15T08:00:00Z', current: '00011', columns: { Id: '1' } },
        { id: 2, created: '2026-10-15T08:01:00Z', current: '00022', columns: { Id: '2' } },
        { id: 3, created: '2026-10-15T08:02:00Z', current: '00033', columns: { Id: '3' } }
      ],
      formula: '{counter:4}{Id}{counter}',
      resetPeriod: 'None',
      padLength: 0,
      now
    });
    expect(plan.plans.map((item) => item.code)).toEqual(['0001', '0002', '0003']);
  });

  it('pulls a counter back when it has moved past the items', () => {
    expect(counterIsCurrent(5, now.toISOString(), 3, 'None', now)).toBe(false);
    expect(counterIsCurrent(3, now.toISOString(), 3, 'None', now)).toBe(true);
    expect(counterIsCurrent(3, '2026-09-01T00:00:00Z', 3, 'Month', now)).toBe(false);
    expect(counterIsCurrent(9, '2026-09-01T00:00:00Z', 0, 'Month', now)).toBe(true);
  });
});