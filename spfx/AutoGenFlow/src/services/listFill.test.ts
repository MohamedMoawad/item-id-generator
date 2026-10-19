import { isBlankNumber, oldestFirst, readColumnValues } from './listFillValues';

describe('list fill', () => {
  it('treats an empty number as blank and reads formula columns', () => {
    expect(isBlankNumber(undefined)).toBe(true);
    expect(isBlankNumber(JSON.parse('null'))).toBe(true);
    expect(isBlankNumber('  ')).toBe(true);
    expect(isBlankNumber('REQ-00001')).toBe(false);
    expect(readColumnValues({
      Title: 'Desk',
      Department: { Title: 'Finance' }
    }, ['Title', 'Department.title'])).toEqual({
      Title: 'Desk',
      'Department.title': 'Finance'
    });
    expect(oldestFirst([{ Id: 9 }, { Id: 8 }])).toEqual([{ Id: 8 }, { Id: 9 }]);
  });
});