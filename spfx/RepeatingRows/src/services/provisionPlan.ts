import { IRepeatingColumn, IRepeatingRowsConfig } from '../models/repeatingRows';

export interface IFieldSpec {
  title: string;
  internalName: string;
  kind: 'text' | 'note' | 'number' | 'currency' | 'choice' | 'date' | 'boolean' | 'lookup' | 'person' | 'taxonomy';
  choices?: string[];
  lookupListId?: string;
  lookupField?: string;
  indexed?: boolean;
  cascadeDelete?: boolean;
}

export function childFieldSpecs(config: IRepeatingRowsConfig, parentListId: string): IFieldSpec[] {
  const relationship: IFieldSpec = config.relationshipKind === 'number'
    ? {
      title: config.relationshipField,
      internalName: config.relationshipField,
      kind: 'number',
      indexed: true
    }
    : {
      title: config.relationshipField,
      internalName: config.relationshipField,
      kind: 'lookup',
      lookupListId: parentListId,
      lookupField: 'Title',
      indexed: true,
      cascadeDelete: true
    };
  const columns = config.columns
    .filter((column) => column.type !== 'attachment' && !column.formula)
    .map(columnSpec);
  return [
    relationship,
    { title: 'SortOrder', internalName: 'SortOrder', kind: 'number', indexed: true },
    ...columns
  ];
}

function columnSpec(column: IRepeatingColumn): IFieldSpec {
  if (column.type === 'taxonomy') {
    return { title: column.title, internalName: column.internalName, kind: 'taxonomy' };
  }
  return {
    title: column.title,
    internalName: column.internalName,
    kind: column.type as IFieldSpec['kind'],
    choices: column.choices,
    lookupField: column.lookupTextField || 'Title'
  };
}
