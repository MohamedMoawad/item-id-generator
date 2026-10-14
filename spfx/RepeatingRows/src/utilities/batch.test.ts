import { buildBatchPayload, chunkOperations, readBatchStatuses, visibleRange } from './batch';

describe('batching', () => {
  it('splits work into fixed chunks and windows the visible rows', () => {
    expect(chunkOperations([1, 2, 3, 4, 5], 2)).toEqual([[1, 2], [3, 4], [5]]);
    expect(visibleRange(100, 720, 480, 72).start).toBeGreaterThan(0);
  });

  it('builds a changeset and reads the status lines', () => {
    const payload = buildBatchPayload([
      { method: 'MERGE', url: 'https://contoso.sharepoint.com/sites/a/_api/web/lists/items(1)', body: '{"SortOrder":2}', etag: '"1"' },
      { method: 'DELETE', url: 'https://contoso.sharepoint.com/sites/a/_api/web/lists/items(2)' }
    ], 'abc');
    expect(payload.contentType).toContain('batch_abc');
    expect(payload.body).toContain('X-HTTP-Method: MERGE');
    expect(payload.body).toContain('DELETE https://contoso.sharepoint.com');
    expect(readBatchStatuses('HTTP/1.1 204 No Content\r\nHTTP/1.1 200 OK')).toEqual([204, 200]);
  });
});
