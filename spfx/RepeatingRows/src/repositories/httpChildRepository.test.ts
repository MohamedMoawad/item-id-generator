import { defaultRepeatingConfig } from '../models/repeatingRows';
import { IHttpResult, ISharePointGateway } from '../services/httpTypes';
import { HttpChildRepository } from './httpChildRepository';

class FakeGateway implements ISharePointGateway {
  public webUrl = 'https://contoso.sharepoint.com/sites/demo';
  public calls: { method: string; url: string; body?: string }[] = [];

  public get(url: string): Promise<IHttpResult> {
    this.calls.push({ method: 'GET', url });
    const item = {
      Id: 8,
      SortOrder: 1,
      ParentIDId: 3,
      Description: 'Item 1',
      Quantity: 2,
      UnitPrice: 100,
      Amount: 200
    };
    return Promise.resolve({ status: 200, text: JSON.stringify({ value: [item] }) });
  }

  public send(method: 'POST' | 'MERGE' | 'DELETE', url: string, body?: string): Promise<IHttpResult> {
    this.calls.push({ method, url, body });
    return Promise.resolve({ status: 201, text: JSON.stringify({ Id: 15 }) });
  }

  public batch(body: string): Promise<IHttpResult> {
    this.calls.push({ method: 'BATCH', url: `${this.webUrl}/_api/$batch`, body });
    return Promise.resolve({ status: 200, text: 'HTTP/1.1 204 No Content\r\nHTTP/1.1 200 OK' });
  }

  public postBinary(url: string): Promise<IHttpResult> {
    this.calls.push({ method: 'BINARY', url });
    return Promise.resolve({ status: 200, text: '' });
  }
}

describe('HttpChildRepository', () => {
  it('pages by parent and batches the edit and delete', async () => {
    const gateway = new FakeGateway();
    const config = defaultRepeatingConfig('Purchase Request Details');
    const repository = new HttpChildRepository(gateway, config);
    const page = await repository.loadPage(3, 0);
    expect(page.rows[0].values.Description).toBe('Item 1');
    expect(page.rows[0].values.Amount).toBe('200');
    expect(gateway.calls[0].url).toContain('ParentIDId%20eq%203');

    await repository.apply({
      create: [{ key: 'new', parentId: 3, sortOrder: 2, values: { Description: 'Item 2', Quantity: '1', UnitPrice: '50', Amount: '50' } }],
      update: [{ key: 'item-8', id: 8, parentId: 3, sortOrder: 1, values: { Description: 'Item 1', Quantity: '3', UnitPrice: '100', Amount: '300' } }],
      deleteIds: [4]
    });
    expect(gateway.calls.some((call) => call.method === 'POST')).toBe(true);
    const batch = gateway.calls.filter((call) => call.method === 'BATCH')[0];
    expect(batch.body || '').toContain('items(8)');
    expect(batch.body || '').toContain('items(4)');
  });
});
