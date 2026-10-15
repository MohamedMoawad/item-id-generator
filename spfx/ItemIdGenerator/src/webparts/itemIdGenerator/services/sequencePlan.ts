import { applyFormula, periodKeyFor, periodRolledOver } from './formula';

export interface ISequenceItem {
  id: number;
  created?: string;
  current: string;
  columns?: { [token: string]: string };
}

export interface IPlannedNumber {
  id: number;
  sequence: number;
  code: string;
  changed: boolean;
}

export interface IContinuousPlan {
  plans: IPlannedNumber[];
  currentCount: number;
  lastResetDate: string;
}

export function planContinuousNumbers(input: {
  items: ISequenceItem[];
  formula: string;
  resetPeriod: string;
  padLength: number;
  now: Date;
}): IContinuousPlan {
  const ordered = input.items.slice().sort((left, right) => left.id - right.id);
  const currentKey = periodKeyFor(input.resetPeriod, input.now);
  const plans: IPlannedNumber[] = [];
  let sequence = 0;
  for (let index = 0; index < ordered.length; index += 1) {
    const item = ordered[index];
    const when = createdDate(item.created, input.now);
    if (periodKeyFor(input.resetPeriod, when) !== currentKey) {
      continue;
    }
    sequence += 1;
    const code = applyFormula(input.formula, sequence, when, input.padLength, item.columns);
    plans.push({
      id: item.id,
      sequence,
      code,
      changed: (item.current || '').trim() !== code
    });
  }
  return {
    plans,
    currentCount: sequence,
    lastResetDate: input.now.toISOString()
  };
}

export function counterIsCurrent(
  currentCount: number,
  lastResetDate: string | undefined,
  plannedCount: number,
  resetPeriod: string,
  now: Date
): boolean {
  if (plannedCount <= 0) {
    return true;
  }
  if (currentCount !== plannedCount) {
    return false;
  }
  return !periodRolledOver(resetPeriod, lastResetDate, now);
}

function createdDate(value: string | undefined, now: Date): Date {
  if (!value) {
    return now;
  }
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? now : parsed;
}
