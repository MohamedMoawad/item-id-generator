import { defaultRepeatingConfig } from '../models/repeatingRows';
import { MemoryChildRepository, MemoryParentRepository } from '../repositories/memoryRepositories';
import { IRepeatingStore } from '../repositories/contracts';
import { saveRepeatingItem } from './repeatingRowsService';
import { RepeatingRowsError } from '../models/repeatingRows';

describe('saveRepeatingItem', () => {
  it('creates the parent first and then the child rows', async () => {
    const parent = new MemoryParentRepository();
    const children = new MemoryChildRepository();
    const store: IRepeatingStore = { parent, children, config: defaultRepeatingConfig('Purchase Request Details') };
    const result = await saveRepeatingItem(store, {
      parentValues: { Title: 'Chairs' },
      rows: [
        { key: 'a', parentId: 0, sortOrder: 1, values: { Description: 'Item 1', Quantity: '2', UnitPrice: '100', Amount: '200' } },
        { key: 'b', parentId: 0, sortOrder: 2, values: { Description: 'Item 2', Quantity: '1', UnitPrice: '50', Amount: '50' } }
      ],
      originalRows: [],
      attachments: [{ rowKey: 'a', fileName: 'quote.pdf', content: new ArrayBuffer(4) }]
    });
    expect(result.parentId).toBe(1);
    expect(parent.items[0].values.RepeatingRows).toBe('2');
    expect(children.rows.map((row) => row.values.Description)).toEqual(['Item 1', 'Item 2']);
    expect(children.attachments).toEqual([{ itemId: result.created[0].id, fileName: 'quote.pdf' }]);
  });

  it('does not write when a required cell is empty', async () => {
    const store: IRepeatingStore = {
      parent: new MemoryParentRepository(),
      children: new MemoryChildRepository(),
      config: defaultRepeatingConfig('Details')
    };
    await expect(saveRepeatingItem(store, {
      parentValues: { Title: 'Blank' },
      rows: [{ key: 'a', parentId: 0, sortOrder: 1, values: { Description: '', Quantity: '1', UnitPrice: '1' } }],
      originalRows: [],
      attachments: []
    })).rejects.toBeInstanceOf(RepeatingRowsError);
  });
});
