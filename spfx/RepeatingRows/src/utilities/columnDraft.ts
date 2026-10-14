import { IRepeatingColumn, RepeatingColumnType } from '../models/repeatingRows';

export const COLUMN_TYPES: { value: RepeatingColumnType; label: string }[] = [
  { value: 'text', label: 'Single line of text' },
  { value: 'note', label: 'Multiple lines of text' },
  { value: 'number', label: 'Number' },
  { value: 'currency', label: 'Currency' },
  { value: 'choice', label: 'Choice' },
  { value: 'date', label: 'Date' },
  { value: 'boolean', label: 'Yes/No' },
  { value: 'lookup', label: 'Lookup' },
  { value: 'person', label: 'Person' },
  { value: 'taxonomy', label: 'Managed metadata (label and guid)' },
  { value: 'attachment', label: 'Attachment' }
];

export function blankColumn(index: number): IRepeatingColumn {
  return {
    internalName: `Column${index}`,
    title: `Column ${index}`,
    type: 'text',
    required: false
  };
}

export function moveColumn(columns: IRepeatingColumn[], fromIndex: number, toIndex: number): IRepeatingColumn[] {
  if (fromIndex < 0 || toIndex < 0 || fromIndex >= columns.length || toIndex >= columns.length || fromIndex === toIndex) {
    return columns;
  }
  const next = columns.slice();
  const [moved] = next.splice(fromIndex, 1);
  next.splice(toIndex, 0, moved);
  return next;
}

export function cleanColumn(column: IRepeatingColumn): IRepeatingColumn {
  const next: IRepeatingColumn = {
    internalName: column.internalName.trim(),
    title: (column.title || column.internalName).trim(),
    type: column.type,
    required: column.required
  };
  assignText(next, 'defaultValue', column.defaultValue);
  assignText(next, 'lookupListTitle', column.lookupListTitle);
  assignText(next, 'lookupTextField', column.lookupTextField);
  assignText(next, 'lookupGroupField', column.lookupGroupField);
  assignText(next, 'cascadeFrom', column.cascadeFrom);
  assignText(next, 'formula', column.formula);
  assignText(next, 'visibleWhen', column.visibleWhen);
  if (column.type === 'choice' && column.choices && column.choices.length > 0) {
    next.choices = column.choices.map((choice) => choice.trim()).filter((choice) => choice.length > 0);
  }
  if (next.formula) {
    next.readOnly = true;
  }
  return next;
}

function assignText(column: IRepeatingColumn, key: 'defaultValue' | 'lookupListTitle' | 'lookupTextField' | 'lookupGroupField' | 'cascadeFrom' | 'formula' | 'visibleWhen', value: string | undefined): void {
  if (value && value.trim()) {
    column[key] = value.trim();
  }
}

export function choicesText(column: IRepeatingColumn): string {
  return (column.choices || []).join(', ');
}

export function withChoices(column: IRepeatingColumn, text: string): IRepeatingColumn {
  const choices = text.split(',').map((choice) => choice.trim()).filter((choice) => choice.length > 0);
  const next: IRepeatingColumn = { ...column };
  if (choices.length > 0) {
    next.choices = choices;
  } else {
    delete next.choices;
  }
  return next;
}
