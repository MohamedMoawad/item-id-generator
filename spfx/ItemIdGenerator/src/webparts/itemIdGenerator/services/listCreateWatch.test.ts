import { createdItemId, isNewItemRequest } from './listCreateWatch';

const listGuid = 'aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee';
const listPath = '/sites/ops/Lists/Requests';

describe('list create watch', () => {
  it('recognizes a new item save for this list', () => {
    expect(isNewItemRequest(
      `https://contoso.sharepoint.com/sites/ops/_api/web/lists(guid'${listGuid}')/items`,
      'POST',
      listGuid,
      listPath
    )).toBe(true);
    expect(isNewItemRequest(
      "https://contoso.sharepoint.com/sites/ops/_api/web/GetList(@a1)/AddValidateUpdateItemUsingPath?@a1='%2Fsites%2Fops%2FLists%2FRequests'",
      'POST',
      listGuid,
      listPath
    )).toBe(true);
    expect(isNewItemRequest(
      `https://contoso.sharepoint.com/sites/ops/_api/web/lists(guid'${listGuid}')/items(12)`,
      'POST',
      listGuid,
      listPath
    )).toBe(false);
  });

  it('reads the new item id from a form save', () => {
    expect(createdItemId({ Id: 12 })).toBe(12);
    expect(createdItemId({
      value: [
        { FieldName: 'Title', FieldValue: 'Desk' },
        { FieldName: 'Id', FieldValue: '18' }
      ]
    })).toBe(18);
    expect(createdItemId({ Title: 'Desk' })).toBeUndefined();
  });
});
