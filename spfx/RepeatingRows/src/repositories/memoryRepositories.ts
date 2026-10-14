import { IChildPage, IChildRow, IParentValues, IPendingAttachment, ISavePlan, RepeatingRowsError } from '../models/repeatingRows';
import { IChildRepository, ICreatedChild, IParentRecord, IParentRepository } from './contracts';

export class MemoryParentRepository implements IParentRepository {
  public items: { id: number; etag: string; values: IParentValues }[] = [];
  private nextId = 1;

  public load(itemId: number): Promise<IParentRecord> {
    const found = this.items.filter((item) => item.id === itemId)[0];
    if (!found) {
      return Promise.reject(new RepeatingRowsError('The list or item was not found.', 404));
    }
    return Promise.resolve({ id: found.id, etag: found.etag, values: { ...found.values }, item: { ...found.values } });
  }

  public create(values: IParentValues): Promise<number> {
    const id = this.nextId;
    this.nextId += 1;
    this.items.push({ id, etag: '1', values: { ...values } });
    return Promise.resolve(id);
  }

  public update(itemId: number, values: IParentValues, etag?: string): Promise<void> {
    const found = this.items.filter((item) => item.id === itemId)[0];
    if (!found) {
      return Promise.reject(new RepeatingRowsError('The list or item was not found.', 404));
    }
    if (etag && found.etag !== etag) {
      return Promise.reject(new RepeatingRowsError('Someone else changed this item. Reload the form and try again.', 412));
    }
    found.values = { ...found.values, ...values };
    found.etag = String(Number(found.etag) + 1);
    return Promise.resolve();
  }
}

export class MemoryChildRepository implements IChildRepository {
  public rows: IChildRow[] = [];
  public attachments: { itemId: number; fileName: string }[] = [];
  private nextId = 1;

  public loadPage(parentId: number, afterSortOrder: number): Promise<IChildPage> {
    const matched = this.rows
      .filter((row) => row.parentId === parentId && row.sortOrder > afterSortOrder)
      .sort((left, right) => left.sortOrder - right.sortOrder);
    return Promise.resolve({
      rows: matched.map((row) => ({ ...row, values: { ...row.values } })),
      hasMore: false
    });
  }

  public apply(plan: ISavePlan): Promise<ICreatedChild[]> {
    const created: ICreatedChild[] = [];
    plan.create.forEach((row) => {
      const id = this.nextId;
      this.nextId += 1;
      this.rows.push({ ...row, id, values: { ...row.values } });
      created.push({ key: row.key, id });
    });
    plan.update.forEach((row) => {
      const index = this.rows.findIndex((item) => item.id === row.id);
      if (index >= 0) {
        this.rows[index] = { ...row, values: { ...row.values } };
      }
    });
    plan.deleteIds.forEach((id) => {
      this.rows = this.rows.filter((row) => row.id !== id);
    });
    return Promise.resolve(created);
  }

  public addAttachments(itemId: number, files: IPendingAttachment[]): Promise<void> {
    files.forEach((file) => {
      this.attachments.push({ itemId, fileName: file.fileName });
    });
    return Promise.resolve();
  }
}
