import { IFieldSpec } from '../services/provisionPlan';
import { MARKER_FIELD_NAME } from '../models/repeatingRows';

export function fieldSchemaXml(spec: IFieldSpec): string {
  const name = xmlEscape(spec.internalName);
  const title = xmlEscape(spec.title);
  const indexed = spec.indexed ? ' Indexed="TRUE"' : '';
  if (spec.kind === 'note' || spec.kind === 'taxonomy') {
    return `<Field Type="Note" Name="${name}" StaticName="${name}" DisplayName="${title}" NumLines="4" />`;
  }
  if (spec.kind === 'number') {
    return `<Field Type="Number" Name="${name}" StaticName="${name}" DisplayName="${title}"${indexed} />`;
  }
  if (spec.kind === 'currency') {
    return `<Field Type="Currency" Name="${name}" StaticName="${name}" DisplayName="${title}" />`;
  }
  if (spec.kind === 'date') {
    return `<Field Type="DateTime" Name="${name}" StaticName="${name}" DisplayName="${title}" Format="DateOnly" />`;
  }
  if (spec.kind === 'boolean') {
    return `<Field Type="Boolean" Name="${name}" StaticName="${name}" DisplayName="${title}" />`;
  }
  if (spec.kind === 'choice') {
    const choices = (spec.choices || []).map((choice) => `<CHOICE>${xmlEscape(choice)}</CHOICE>`).join('');
    return `<Field Type="Choice" Name="${name}" StaticName="${name}" DisplayName="${title}"><CHOICES>${choices}</CHOICES></Field>`;
  }
  if (spec.kind === 'person') {
    return `<Field Type="User" Name="${name}" StaticName="${name}" DisplayName="${title}" UserSelectionMode="PeopleOnly" />`;
  }
  if (spec.kind === 'lookup') {
    const listId = (spec.lookupListId || '').replace(/[{}]/g, '');
    const show = xmlEscape(spec.lookupField || 'Title');
    const cascade = spec.cascadeDelete ? ' RelationshipDeleteBehavior="Cascade"' : '';
    return `<Field Type="Lookup" Name="${name}" StaticName="${name}" DisplayName="${title}" List="{${listId}}" ShowField="${show}"${indexed}${cascade} />`;
  }
  return `<Field Type="Text" Name="${name}" StaticName="${name}" DisplayName="${title}" />`;
}

export function markerFieldSchema(): string {
  return `<Field Type="Text" Name="${MARKER_FIELD_NAME}" StaticName="${MARKER_FIELD_NAME}" DisplayName="Line Items" Group="Repeating Rows" ShowInNewForm="FALSE" ShowInEditForm="FALSE" ShowInDisplayForm="FALSE" />`;
}

export function xmlEscape(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
