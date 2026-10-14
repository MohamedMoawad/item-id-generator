import { IChildPage, IChildRow, IPendingAttachment, IRepeatingRowsConfig, ISavePlan, PAGE_SIZE, RepeatingRowsError } from '../models/repeatingRows';
import { IChildRepository, ICreatedChild } from './contracts';
import { assertOk, ISharePointGateway, odataTitle, parseJson, readId, readRows } from '../services/httpTypes';
import { BATCH_CHUNK, buildBatchPayload, chunkOperations, IBatchRequest, readBatchStatuses } from '../utilities/batch';
import { messageFromStatus } from '../utilities/errors';
import { childQuery, itemToRow, relationshipIdName, rowPayload } from '../utilities/fieldValues';

export class HttpChildRepository implements IChildRepository {
  public constructor(
    private readonly gateway: ISharePointGateway,
    private readonly config: IRepeatingRowsConfig
  ) {}

  public async loadPage(parentId: number, afterSortOrder: number): Promise<IChildPage> {
    const query = childQuery(this.config);
    const relation = relationshipIdName(this.config);
    const filter = afterSortOrder > 0
      ? `${relation} eq ${parentId} and SortOrder gt ${afterSortOrder}`
      : `${relation} eq ${parentId}`;
    const expand = query.expand.length > 0 ? `&$expand=${query.expand.join(',')}` : '';
    const url = `${this.itemsUrl()}?$select=${query.select.join(',')}&$filter=${encodeURIComponent(filter)}&$orderby=SortOrder,Id&$top=${PAGE_SIZE}${expand}`;
    const result = await this.gateway.get(url);
    assertOk(result);
    const items = readRows<{ [key: string]: unknown }>(result.text);
    return {
      rows: items.map((item) => itemToRow(item, this.config)),
      hasMore: items.length === PAGE_SIZE
    };
  }

  public async apply(plan: ISavePlan): Promise<ICreatedChild[]> {
    const created = await this.createRows(plan.create);
    const changes: IBatchRequest[] = [];
    plan.update.forEach((row) => {
      if (row.id === undefined) {
        return;
      }
      changes.push({
        method: 'MERGE',
        url: `${this.itemsUrl()}(${row.id})`,
        body: JSON.stringify(rowPayload(row, this.config)),
        etag: row.etag || '*'
      });
    });
    plan.deleteIds.forEach((id) => {
      changes.push({ method: 'DELETE', url: `${this.itemsUrl()}(${id})` });
    });
    const groups = chunkOperations(changes, BATCH_CHUNK);
    for (let index = 0; index < groups.length; index += 1) {
      await this.postBatch(groups[index], index);
    }
    return created;
  }

  public async addAttachments(itemId: number, files: IPendingAttachment[]): Promise<void> {
    for (let index = 0; index < files.length; index += 1) {
      const file = files[index];
      const safeName = encodeURIComponent(file.fileName).replace(/'/g, "''");
      const result = await this.gateway.postBinary(
        `${this.itemsUrl()}(${itemId})/AttachmentFiles/add(FileName='${safeName}')`,
        file.content
      );
      assertOk(result);
    }
  }

  private async createRows(rows: IChildRow[]): Promise<ICreatedChild[]> {
    const created: ICreatedChild[] = [];
    const groups = chunkOperations(rows, BATCH_CHUNK);
    for (let index = 0; index < groups.length; index += 1) {
      const group = groups[index];
      const ids = await Promise.all(group.map((row) => this.createOne(row)));
      ids.forEach((id, rowIndex) => created.push({ key: group[rowIndex].key, id }));
    }
    return created;
  }

  private async createOne(row: IChildRow): Promise<number> {
    const result = await this.gateway.send('POST', this.itemsUrl(), JSON.stringify(rowPayload(row, this.config)));
    assertOk(result);
    return readId(result.text);
  }

  private async postBatch(requests: IBatchRequest[], index: number): Promise<void> {
    if (requests.length === 0) {
      return;
    }
    const payload = buildBatchPayload(requests, `rows${index}${Date.now().toString(36)}`);
    const result = await this.gateway.batch(payload.body, payload.contentType);
    if (result.status >= 400) {
      throw new RepeatingRowsError(messageFromStatus(result.status, result.text), result.status);
    }
    const statuses = readBatchStatuses(result.text);
    const failed = statuses.filter((status) => status >= 400);
    if (failed.length > 0) {
      throw new RepeatingRowsError(messageFromStatus(failed[0], result.text), failed[0]);
    }
  }

  private itemsUrl(): string {
    return `${this.gateway.webUrl}/_api/web/lists/getbytitle('${odataTitle(this.config.childListTitle)}')/items`;
  }
}

export function attachmentNames(text: string): string[] {
  const payload = parseJson<{ value?: { FileName?: string }[] }>(text);
  return (payload.value || []).map((file) => file.FileName || '').filter((name) => name.length > 0);
}
