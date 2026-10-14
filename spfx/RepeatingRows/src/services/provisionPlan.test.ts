import { defaultRepeatingConfig } from '../models/repeatingRows';
import { childFieldSpecs } from './provisionPlan';
import { fieldSchemaXml, markerFieldSchema } from '../utilities/fieldSchema';

describe('provisioning schema', () => {
  it('creates a cascading lookup to the parent and indexes the sort order', () => {
    const config = defaultRepeatingConfig('Purchase Request Details');
    const specs = childFieldSpecs(config, 'parent-guid');
    const parent = specs[0];
    expect(parent.kind).toBe('lookup');
    expect(parent.cascadeDelete).toBe(true);
    expect(fieldSchemaXml(parent)).toContain('RelationshipDeleteBehavior="Cascade"');
    expect(fieldSchemaXml(parent)).toContain('List="{parent-guid}"');
    expect(fieldSchemaXml(specs[1])).toContain('Indexed="TRUE"');
    expect(markerFieldSchema()).toContain('ShowInNewForm="FALSE"');
  });
});
