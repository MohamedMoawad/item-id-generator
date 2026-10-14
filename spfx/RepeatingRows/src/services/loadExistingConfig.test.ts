import { IHttpResult, ISharePointGateway } from './httpTypes';
import { loadExistingConfig } from './loadExistingConfig';

class FakeGateway implements ISharePointGateway {
  public constructor(private readonly status: number, private readonly text: string) {}
  public webUrl = 'https://contoso.sharepoint.com/sites/demo';
  public get(url: string): Promise<IHttpResult> {
    expect(url).toContain('RepeatingRows');
    return Promise.resolve({ status: this.status, text: this.text });
  }
  public send(): Promise<IHttpResult> {
    return Promise.resolve({ status: 200, text: '' });
  }
  public batch(): Promise<IHttpResult> {
    return Promise.resolve({ status: 200, text: '' });
  }
  public postBinary(): Promise<IHttpResult> {
    return Promise.resolve({ status: 200, text: '' });
  }
}

describe('loadExistingConfig', () => {
  it('reads the saved column configuration', async () => {
    const properties = JSON.stringify({
      childListTitle: 'Purchase Request Details',
      relationshipField: 'ParentID',
      minRows: 2,
      columns: [{ internalName: 'Description', title: 'Description', type: 'text', required: true }]
    });
    const config = await loadExistingConfig(
      new FakeGateway(200, JSON.stringify({ ClientSideComponentProperties: properties })),
      'parent-guid',
      'Fallback'
    );
    expect(config && config.childListTitle).toBe('Purchase Request Details');
    expect(config && config.minRows).toBe(2);
  });

  it('returns nothing when the column has not been created', async () => {
    const config = await loadExistingConfig(new FakeGateway(404, ''), 'parent-guid', 'Fallback');
    expect(config).toBeUndefined();
  });
});
