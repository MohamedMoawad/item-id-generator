import { RepeatingRowsError } from '../models/repeatingRows';
import { parseRepeatingConfig } from './configParser';

describe('config parser', () => {
  it('reads the administrator settings', () => {
    const config = parseRepeatingConfig({
      childListTitle: 'Purchase Request Details',
      relationshipField: 'ParentID',
      minRows: 1,
      maxRows: 20,
      allowDelete: false,
      columns: [{ internalName: 'Description', title: 'Description', type: 'text', required: true }]
    }, 'Fallback');
    expect(config.childListTitle).toBe('Purchase Request Details');
    expect(config.allowDelete).toBe(false);
    expect(config.allowEdit).toBe(true);
    expect(config.relationshipKind).toBe('lookup');
  });

  it('rejects an internal name that SharePoint cannot store', () => {
    expect(() => parseRepeatingConfig({
      childListTitle: 'Details',
      columns: [{ internalName: 'Unit Price', title: 'Unit Price', type: 'text', required: false }]
    }, 'Details')).toThrow(RepeatingRowsError);
  });
});
