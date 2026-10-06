import { buildGenerateItemIdBody } from './generateItemIdClient';
import { buildListItemMerge } from './updateListItemField';

describe('item id request contract', () => {
  it('posts listName, condition, and itemId', () => {
    const body = buildGenerateItemIdBody({
      functionUrl: 'https://contoso.example/api/GenerateItemId',
      listName: ' Requests ',
      condition: ' VIP ',
      itemId: 15
    });

    expect(JSON.parse(body)).toEqual({
      listName: 'Requests',
      condition: 'VIP',
      itemId: 15
    });
  });

  it('builds a MERGE url from the list title and field internal name', () => {
    const merge = buildListItemMerge({
      webAbsoluteUrl: 'https://contoso.sharepoint.com/sites/ops/',
      listName: "Team's Requests",
      itemId: 15,
      targetFieldInternalName: 'GeneratedItemId',
      generatedId: 'VIP-Requests-1700000000000'
    });

    expect(merge.url).toBe(
      "https://contoso.sharepoint.com/sites/ops/_api/web/lists/getbytitle('Team''s Requests')/items(15)"
    );
    expect(JSON.parse(merge.body)).toEqual({
      GeneratedItemId: 'VIP-Requests-1700000000000'
    });
  });

  it('rejects a field internal name that is not a SharePoint identifier', () => {
    expect(() => buildListItemMerge({
      webAbsoluteUrl: 'https://contoso.sharepoint.com/sites/ops',
      listName: 'Requests',
      itemId: 15,
      targetFieldInternalName: 'Generated ID',
      generatedId: 'Requests-1'
    })).toThrow(/internal name/);
  });
});
