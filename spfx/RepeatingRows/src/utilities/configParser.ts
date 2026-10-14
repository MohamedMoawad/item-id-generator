import {
  defaultRepeatingConfig,
  IRepeatingColumn,
  IRepeatingRowsConfig,
  RepeatingColumnType,
  RepeatingRowsError
} from '../models/repeatingRows';

const TYPES: RepeatingColumnType[] = [
  'text', 'note', 'number', 'currency', 'choice', 'date', 'boolean', 'lookup', 'person', 'taxonomy', 'attachment'
];

export function parseRepeatingConfig(raw: unknown, fallbackChildList: string): IRepeatingRowsConfig {
  const record = asRecord(raw);
  const fallback = defaultRepeatingConfig(fallbackChildList);
  const columns = Array.isArray(record.columns) ? record.columns.map(parseColumn).filter(isColumn) : fallback.columns;
  if (columns.length === 0) {
    throw new RepeatingRowsError('Add at least one column.');
  }
  const childListTitle = text(record.childListTitle) || fallback.childListTitle;
  if (!childListTitle.trim()) {
    throw new RepeatingRowsError('Enter the child list name.');
  }
  return {
    childListTitle: childListTitle.trim(),
    relationshipField: fieldName(text(record.relationshipField) || fallback.relationshipField),
    relationshipKind: text(record.relationshipKind) === 'number' ? 'number' : 'lookup',
    columns,
    minRows: whole(record.minRows, fallback.minRows),
    maxRows: whole(record.maxRows, fallback.maxRows),
    allowDelete: record.allowDelete === undefined ? fallback.allowDelete : Boolean(record.allowDelete),
    allowEdit: record.allowEdit === undefined ? fallback.allowEdit : Boolean(record.allowEdit)
  };
}

export function serializeConfig(config: IRepeatingRowsConfig): string {
  return JSON.stringify(config);
}

function parseColumn(value: unknown): IRepeatingColumn | undefined {
  const record = asRecord(value);
  const internalName = text(record.internalName);
  const title = text(record.title) || internalName;
  const type = text(record.type) as RepeatingColumnType;
  if (!internalName || TYPES.indexOf(type) < 0) {
    return undefined;
  }
  const column: IRepeatingColumn = {
    internalName: fieldName(internalName),
    title,
    type,
    required: Boolean(record.required)
  };
  if (text(record.defaultValue)) {
    column.defaultValue = text(record.defaultValue);
  }
  if (Array.isArray(record.choices)) {
    column.choices = record.choices.map((choice) => String(choice)).filter((choice) => choice.length > 0);
  }
  if (text(record.lookupListTitle)) {
    column.lookupListTitle = text(record.lookupListTitle);
  }
  if (text(record.lookupTextField)) {
    column.lookupTextField = text(record.lookupTextField);
  }
  if (text(record.cascadeFrom)) {
    column.cascadeFrom = text(record.cascadeFrom);
  }
  if (text(record.lookupGroupField)) {
    column.lookupGroupField = text(record.lookupGroupField);
  }
  if (text(record.formula)) {
    column.formula = text(record.formula);
  }
  if (text(record.visibleWhen)) {
    column.visibleWhen = text(record.visibleWhen);
  }
  if (record.readOnly === true) {
    column.readOnly = true;
  }
  return column;
}

function isColumn(column: IRepeatingColumn | undefined): column is IRepeatingColumn {
  return !!column;
}

function fieldName(value: string): string {
  const name = value.trim();
  if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) {
    throw new RepeatingRowsError('Column internal names use letters, numbers, and underscores.');
  }
  return name;
}

function whole(value: unknown, fallback: number): number {
  const numeric = Number(value);
  if (!Number.isInteger(numeric) || numeric < 0) {
    return fallback;
  }
  return numeric;
}

function text(value: unknown): string {
  return typeof value === 'string' ? value : '';
}

function asRecord(value: unknown): { [key: string]: unknown } {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    return {};
  }
  return value as { [key: string]: unknown };
}
