import { buildItemQuery } from './itemQuery';

describe('item query', () => {
  it('asks for blank items in creation order and expands lookup titles', () => {
    expect(buildItemQuery(
      '{AABBCCDD-1111-2222-3333-444455556666}',
      'autogenflow',
      'Title,Department.title'
    )).toBe("_api/web/lists(guid'aabbccdd-1111-2222-3333-444455556666')/items?$select=Id,Created,autogenflow,Title,Department/Title&$filter=autogenflow+eq+null&$orderby=Id+asc&$top=100&$expand=Department");
  });
});
