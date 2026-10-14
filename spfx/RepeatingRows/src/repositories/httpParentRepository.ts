import { IParentValues, RepeatingRowsError } from '../models/repeatingRows';
import { IParentRecord, IParentRepository } from './contracts';
import { assertOk, ISharePointGateway, parseJson, readId } from '../services/httpTypes';

export class HttpParentRepository implements IParentRepository {
  public constructor(
    private readonly gateway: ISharePointGateway,
    private readonly listId: string,
    private readonly selectFields: string[],
    private readonly expandFields: string[]
  ) {}

  public async load(itemId: number): Promise<IParentRecord> {
    const expand = this.expandFields.length > 0 ? `&$expand=${this.expandFields.join(',')}` : '';
    const select = ['Id', ...this.selectFields].join(',');
    const url = `${this.listUrl()}/items(${itemId})?$select=${select}${expand}`;
    const result = await this.gateway.get(url);
    assertOk(result);
    const item = parseJson<{ [key: string]: unknown }>(result.text);
    const etag = typeof item['odata.etag'] === 'string' ? item['odata.etag'] : undefined;
    return { id: itemId, etag, values: {}, item };
  }

  public async create(values: IParentValues): Promise<number> {
    const result = await this.gateway.send('POST', `${this.listUrl()}/items`, JSON.stringify(values));
    assertOk(result);
    return readId(result.text);
  }

  public async update(itemId: number, values: IParentValues, etag?: string): Promise<void> {
    const result = await this.gateway.send('MERGE', `${this.listUrl()}/items(${itemId})`, JSON.stringify(values), etag);
    if (result.status === 412) {
      throw new RepeatingRowsError('Someone else changed this item. Reload the form and try again.', 412);
    }
    assertOk(result);
  }

  private listUrl(): string {
    return `${this.gateway.webUrl}/_api/web/lists(guid'${this.listId}')`;
  }
}
