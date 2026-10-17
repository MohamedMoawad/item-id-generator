import { buildItemQuery } from './itemQuery';

describe('item query', () => {
  it('selects the created item fields and expands lookup titles', () => {
    expect(buildItemQuery('autogenflow', 'Title,Department.title')).toBe(
      '$select=Id,Created,autogenflow,Title,Department/Title&$expand=Department'
    );
  });
});
