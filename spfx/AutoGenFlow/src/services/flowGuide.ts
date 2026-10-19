import { normalizeGuid } from './contract';

export const FLOW_CONCURRENCY = 1;

export interface IFlowExpression {
  action: string;
  expression: string;
}

export function configFilter(listGuid: string): string {
  return `TargetListGuid eq '${normalizeGuid(listGuid)}' and IsActive eq 1`;
}

export function triggerCondition(numberColumn: string): string {
  return `@empty(triggerOutputs()?['body/${numberColumn.trim()}'])`;
}

export function blankGuardExpression(numberColumn: string): string {
  return `and(empty(outputs('Get_item')?['body/${numberColumn.trim()}']),greater(length(outputs('Get_items')?['body/value']),0))`;
}

export const CONFIG_EXPRESSION = `first(outputs('Get_items')?['body/value'])`;

export const RESET_PERIOD_EXPRESSION = `outputs('Config')?['ResetPeriod']`;

export const LAST_RESET_EXPRESSION = `outputs('Config')?['LastResetDate']`;

export const CURRENT_COUNT_EXPRESSION = `int(coalesce(outputs('Config')?['CurrentCount'], 0))`;

export const COUNTER_WIDTH_EXPRESSION = `int(coalesce(outputs('Config')?['CounterWidth'], 0))`;

export const COUNTER_TOKEN_EXPRESSION = `outputs('Config')?['CounterToken']`;

export const FORMULA_COLUMNS_EXPRESSION = `coalesce(outputs('Config')?['FormulaColumns'], '')`;

export const ROLLED_EXPRESSION = [
  "if(",
  "or(equals(outputs('ResetPeriod'), 'None'), empty(outputs('LastResetDate'))),",
  "false,",
  "not(equals(",
  "if(equals(outputs('ResetPeriod'), 'Day'), formatDateTime(outputs('LastResetDate'), 'yyyy-MM-dd'), if(equals(outputs('ResetPeriod'), 'Month'), formatDateTime(outputs('LastResetDate'), 'yyyy-MM'), if(equals(outputs('ResetPeriod'), 'Year'), formatDateTime(outputs('LastResetDate'), 'yyyy'), ''))),",
  "if(equals(outputs('ResetPeriod'), 'Day'), formatDateTime(triggerOutputs()?['body/Created'], 'yyyy-MM-dd'), if(equals(outputs('ResetPeriod'), 'Month'), formatDateTime(triggerOutputs()?['body/Created'], 'yyyy-MM'), if(equals(outputs('ResetPeriod'), 'Year'), formatDateTime(triggerOutputs()?['body/Created'], 'yyyy'), '')))",
  "))",
  ")"
].join('');

export const NEXT_EXPRESSION = `if(outputs('Rolled'), 1, add(outputs('CurrentCount'), 1))`;

export const PADDED_EXPRESSION = [
  "if(",
  "or(equals(outputs('CounterWidth'), 0), greater(length(string(outputs('Next'))), outputs('CounterWidth'))),",
  "string(outputs('Next')),",
  "substring(concat('000000000000', string(outputs('Next'))), sub(12, sub(outputs('CounterWidth'), length(string(outputs('Next'))))), outputs('CounterWidth'))",
  ")"
].join('');

export const CODE_START_EXPRESSION = [
  "replace(",
  "replace(",
  "replace(",
  "replace(",
  "replace(",
  "replace(",
  "replace(outputs('Config')?['Formula'], outputs('CounterToken'), outputs('Padded')),",
  "'{yyyy}', formatDateTime(triggerOutputs()?['body/Created'], 'yyyy')),",
  "'{yy}', formatDateTime(triggerOutputs()?['body/Created'], 'yy')),",
  "'{MM}', formatDateTime(triggerOutputs()?['body/Created'], 'MM')),",
  "'{dd}', formatDateTime(triggerOutputs()?['body/Created'], 'dd')),",
  "'{HH}', formatDateTime(triggerOutputs()?['body/Created'], 'HH')),",
  "'{mm}', formatDateTime(triggerOutputs()?['body/Created'], 'mm'))"
].join('');

export const COLUMN_REPLACE_EXPRESSION = [
  "replace(",
  "variables('Code'),",
  "concat('{', item(), '}'),",
  "coalesce(",
  "triggerOutputs()?['body']?[first(split(item(), '.'))]?['Value'],",
  "triggerOutputs()?['body']?[first(split(item(), '.'))]?['Title'],",
  "string(triggerOutputs()?['body']?[first(split(item(), '.'))]))",
  ")"
].join('');

export function flowExpressions(numberColumn: string): IFlowExpression[] {
  return [
    { action: 'Condition', expression: blankGuardExpression(numberColumn) },
    { action: 'Config', expression: CONFIG_EXPRESSION },
    { action: 'ResetPeriod', expression: RESET_PERIOD_EXPRESSION },
    { action: 'LastResetDate', expression: LAST_RESET_EXPRESSION },
    { action: 'CurrentCount', expression: CURRENT_COUNT_EXPRESSION },
    { action: 'CounterWidth', expression: COUNTER_WIDTH_EXPRESSION },
    { action: 'CounterToken', expression: COUNTER_TOKEN_EXPRESSION },
    { action: 'FormulaColumns', expression: FORMULA_COLUMNS_EXPRESSION },
    { action: 'Rolled', expression: ROLLED_EXPRESSION },
    { action: 'Next', expression: NEXT_EXPRESSION },
    { action: 'Padded', expression: PADDED_EXPRESSION },
    { action: 'CodeStart', expression: CODE_START_EXPRESSION },
    { action: 'Set Code inside Apply to each', expression: COLUMN_REPLACE_EXPRESSION }
  ];
}

export function flowSteps(): string[] {
  return [
    'Create one scheduled cloud flow for the site collection. Recurrence, every 1 minute. Concurrency control on, degree of parallelism 1.',
    'Get items from AutoGenFlowConfiguration on the site collection. Filter IsActive eq 1 and TargetListGuid ne \'site\'. Top Count 200.',
    'Apply to each rule, concurrency 1. Skip a row whose TargetListGuid is site.',
    'Take the web address from TargetListUrl: the part before /Lists/, or the part before the last slash for a library.',
    'Send an HTTP request to that web. Get the newest 25 items on the list in TargetListGuid, using ItemQuery from the row, ordered by Modified descending.',
    'Apply to each item, oldest of that page first, concurrency 1. Continue when the number column is empty.',
    'Initialize Code, Next, and Previous outside the loops. Previous holds the last committed count for the current rule.',
    'Add one Compose action for each name below and paste its expression. The action name must match exactly.',
    'Set Next to Previous plus 1. Set Code from the formula on the config row.',
    'If FormulaColumns is not empty, Apply to each split of FormulaColumns and set Code with the column expression.',
    'Update the config row first: Current count becomes Next and Last reset date becomes the current time. Then write Code into the number column.',
    'If the item write fails, put Current count back to Previous. Leave the same flow in place for every other list.'
  ];
}
