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
    'Create an automated cloud flow with the SharePoint trigger When an item is created. Choose this list.',
    `Open the trigger settings. Turn concurrency control on and set degree of parallelism to ${FLOW_CONCURRENCY}. Add the trigger condition shown below.`,
    'Add Get item for this same list. Id is the trigger ID. This reads the item again so a second run does not overwrite a number.',
    'Add Get items on the master list site and AutoGenFlowConfiguration. Use the filter query and Top Count 1.',
    'Initialize a string variable named Code and leave it empty. Initialize variable stays outside the condition.',
    'Add a condition with the blank-item expression below. Put every remaining action in the Yes branch.',
    'Add one Compose action for each name below and paste its expression. The action name must match exactly.',
    'Set variable Code to the CodeStart output.',
    'If FormulaColumns is not empty, Apply to each split(outputs(\'FormulaColumns\'), \',\') and set Code with the column expression.',
    'Update the created item. Set the number column to variables(\'Code\').',
    'Update the master row Id outputs(\'Config\')?[\'ID\']. Set Current count to the Next output and Last reset date to Created from the trigger.',
    'Create one flow per list. Clone it and change the trigger list, the trigger condition, and the filter query.'
  ];
}
