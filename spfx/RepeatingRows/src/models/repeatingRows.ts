export const FIELD_CUSTOMIZER_ID = '8e4b1c72-5a90-4d33-b7e1-0c6f2a9d4b18';
export const FORM_CUSTOMIZER_ID = '1d7a6e45-3c28-4f90-a1b6-9e5d0c3f7a64';
export const SETUP_COMMAND_ID = '5b9e2d80-6f14-4c57-8a23-d0e7b1c4f956';

export const MARKER_FIELD_NAME = 'RepeatingRows';
export const PAGE_SIZE = 100;

export type RepeatingColumnType =
  | 'text'
  | 'note'
  | 'number'
  | 'currency'
  | 'choice'
  | 'date'
  | 'boolean'
  | 'lookup'
  | 'person'
  | 'taxonomy'
  | 'attachment';

export type RelationshipKind = 'lookup' | 'number';

export interface IRepeatingColumn {
  internalName: string;
  title: string;
  type: RepeatingColumnType;
  required: boolean;
  defaultValue?: string;
  choices?: string[];
  lookupListTitle?: string;
  lookupTextField?: string;
  /** Another row column. Options are limited to those whose group equals that column's value. */
  cascadeFrom?: string;
  /** Lookup-list column used as the cascade group. */
  lookupGroupField?: string;
  formula?: string;
  /** Show the column when another row column equals this value. Example: Status=Open */
  visibleWhen?: string;
  readOnly?: boolean;
}

export interface IRepeatingRowsConfig {
  childListTitle: string;
  relationshipField: string;
  /** Lookup stores the parent as a real lookup. Number is the fallback when a lookup cannot be created. */
  relationshipKind: RelationshipKind;
  columns: IRepeatingColumn[];
  minRows: number;
  maxRows: number;
  allowDelete: boolean;
  allowEdit: boolean;
}

export interface IChildRow {
  key: string;
  id?: number;
  parentId: number;
  sortOrder: number;
  values: { [internalName: string]: string };
  etag?: string;
}

export interface ILookupOption {
  id: string;
  text: string;
  group?: string;
}

export interface IValidationIssue {
  rowKey: string;
  column?: string;
  message: string;
}

export interface ISavePlan {
  create: IChildRow[];
  update: IChildRow[];
  deleteIds: number[];
}

export interface IParentValues {
  [internalName: string]: string | number | boolean;
}

export interface IParentField {
  internalName: string;
  title: string;
  type: 'text' | 'note' | 'number' | 'currency' | 'choice' | 'date' | 'boolean' | 'person';
  required: boolean;
  choices?: string[];
}

export interface IPendingAttachment {
  rowKey: string;
  fileName: string;
  content: ArrayBuffer;
}

export interface IPersonOption {
  id: string;
  login: string;
  text: string;
}

export interface IChildPage {
  rows: IChildRow[];
  hasMore: boolean;
}

export class RepeatingRowsError extends Error {
  public readonly status?: number;

  public constructor(message: string, status?: number) {
    super(message);
    Object.setPrototypeOf(this, RepeatingRowsError.prototype);
    this.name = 'RepeatingRowsError';
    this.status = status;
  }
}

export function purchaseLineColumns(): IRepeatingColumn[] {
  return [
    { internalName: 'Description', title: 'Description', type: 'text', required: true },
    { internalName: 'Quantity', title: 'Qty', type: 'number', required: true, defaultValue: '1' },
    { internalName: 'UnitPrice', title: 'Unit Price', type: 'currency', required: true, defaultValue: '0' },
    {
      internalName: 'Amount',
      title: 'Amount',
      type: 'currency',
      required: false,
      readOnly: true,
      formula: '{Quantity}*{UnitPrice}'
    }
  ];
}

export function defaultRepeatingConfig(childListTitle: string): IRepeatingRowsConfig {
  return {
    childListTitle,
    relationshipField: 'ParentID',
    relationshipKind: 'lookup',
    columns: purchaseLineColumns(),
    minRows: 1,
    maxRows: 0,
    allowDelete: true,
    allowEdit: true
  };
}
