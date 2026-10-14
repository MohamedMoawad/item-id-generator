import { IChildPage, IChildRow, IParentValues, IPendingAttachment, IRepeatingRowsConfig, ISavePlan } from '../models/repeatingRows';

export interface IParentRecord {
  id: number;
  etag?: string;
  values: IParentValues;
  item: { [key: string]: unknown };
}

export interface ICreatedChild {
  key: string;
  id: number;
}

export interface IParentRepository {
  load(itemId: number): Promise<IParentRecord>;
  create(values: IParentValues): Promise<number>;
  update(itemId: number, values: IParentValues, etag?: string): Promise<void>;
}

export interface IChildRepository {
  loadPage(parentId: number, afterSortOrder: number): Promise<IChildPage>;
  apply(plan: ISavePlan): Promise<ICreatedChild[]>;
  addAttachments(itemId: number, files: IPendingAttachment[]): Promise<void>;
}

export interface IRepeatingStore {
  parent: IParentRepository;
  children: IChildRepository;
  config: IRepeatingRowsConfig;
}

export interface ISaveRequest {
  parentId?: number;
  etag?: string;
  parentValues: IParentValues;
  rows: IChildRow[];
  originalRows: IChildRow[];
  attachments: IPendingAttachment[];
}

export interface ISaveResult {
  parentId: number;
  created: ICreatedChild[];
}
