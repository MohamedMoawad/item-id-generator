import {
  FIELD_CUSTOMIZER_ID,
  FORM_CUSTOMIZER_ID,
  IRepeatingRowsConfig,
  MARKER_FIELD_NAME,
  RepeatingRowsError
} from '../models/repeatingRows';
import { childFieldSpecs, IFieldSpec } from './provisionPlan';
import { fieldSchemaXml, markerFieldSchema } from '../utilities/fieldSchema';
import { serializeConfig } from '../utilities/configParser';
import { assertOk, ISharePointGateway, odataTitle, parseJson, readRows } from './httpTypes';

const fieldHeaders: { [key: string]: string } = {
  Accept: 'application/json;odata.metadata=none',
  'Content-Type': 'application/json;charset=utf-8',
  'odata-version': ''
};

export async function provisionRepeatingRows(
  gateway: ISharePointGateway,
  parentListId: string,
  config: IRepeatingRowsConfig,
  sendField: (url: string, body: string) => Promise<{ status: number; text: string }>
): Promise<void> {
  const childListId = await ensureList(gateway, config.childListTitle);
  await merge(gateway, `${gateway.webUrl}/_api/web/lists(guid'${childListId}')`, JSON.stringify({ EnableVersioning: true }));
  const specs = await resolveLookups(gateway, childFieldSpecs(config, parentListId), config);
  await ensureFields(gateway, childListId, specs, sendField);
  await ensureMarker(gateway, parentListId, serializeConfig(config), sendField);
  await addToDefaultView(gateway, parentListId);
  await registerForms(gateway, parentListId, serializeConfig(config));
}

async function ensureList(gateway: ISharePointGateway, title: string): Promise<string> {
  const existing = await gateway.get(`${gateway.webUrl}/_api/web/lists/getbytitle('${odataTitle(title)}')?$select=Id`);
  if (existing.status === 200) {
    return parseJson<{ Id: string }>(existing.text).Id;
  }
  if (existing.status !== 404) {
    assertOk(existing);
  }
  const created = await gateway.send('POST', `${gateway.webUrl}/_api/web/lists`, JSON.stringify({
    Title: title,
    Description: 'Line items for repeating rows',
    BaseTemplate: 100
  }));
  assertOk(created);
  const id = parseJson<{ Id?: string }>(created.text).Id;
  if (!id) {
    throw new RepeatingRowsError('SharePoint did not return the child list id.');
  }
  return id;
}

async function resolveLookups(
  gateway: ISharePointGateway,
  specs: IFieldSpec[],
  config: IRepeatingRowsConfig
): Promise<IFieldSpec[]> {
  const resolved: IFieldSpec[] = [];
  for (let index = 0; index < specs.length; index += 1) {
    const spec = specs[index];
    const column = config.columns.filter((item) => item.internalName === spec.internalName)[0];
    if (spec.kind === 'lookup' && column && column.lookupListTitle && !spec.lookupListId) {
      const list = await gateway.get(
        `${gateway.webUrl}/_api/web/lists/getbytitle('${odataTitle(column.lookupListTitle)}')?$select=Id`
      );
      assertOk(list);
      resolved.push({ ...spec, lookupListId: parseJson<{ Id: string }>(list.text).Id });
    } else {
      resolved.push(spec);
    }
  }
  return resolved;
}

async function ensureFields(
  gateway: ISharePointGateway,
  listId: string,
  specs: IFieldSpec[],
  sendField: (url: string, body: string) => Promise<{ status: number; text: string }>
): Promise<void> {
  const present = await fieldNames(gateway, listId);
  for (let index = 0; index < specs.length; index += 1) {
    const spec = specs[index];
    if (present.has(spec.internalName.toLowerCase())) {
      continue;
    }
    await createField(gateway, listId, fieldSchemaXml(spec), sendField);
  }
}

async function ensureMarker(
  gateway: ISharePointGateway,
  listId: string,
  properties: string,
  sendField: (url: string, body: string) => Promise<{ status: number; text: string }>
): Promise<void> {
  const present = await fieldNames(gateway, listId);
  if (!present.has(MARKER_FIELD_NAME.toLowerCase())) {
    await createField(gateway, listId, markerFieldSchema(), sendField);
  }
  await merge(
    gateway,
    `${gateway.webUrl}/_api/web/lists(guid'${listId}')/fields/getbyinternalnameortitle('${MARKER_FIELD_NAME}')`,
    JSON.stringify({
      ClientSideComponentId: FIELD_CUSTOMIZER_ID,
      ClientSideComponentProperties: properties,
      Title: 'Line Items'
    })
  );
}

async function addToDefaultView(gateway: ISharePointGateway, listId: string): Promise<void> {
  const view = await gateway.get(`${gateway.webUrl}/_api/web/lists(guid'${listId}')/defaultview?$select=Id`);
  assertOk(view);
  const viewId = parseJson<{ Id: string }>(view.text).Id;
  const added = await gateway.send(
    'POST',
    `${gateway.webUrl}/_api/web/lists(guid'${listId}')/views(guid'${viewId}')/viewfields/addviewfield('${MARKER_FIELD_NAME}')`,
    '{}'
  );
  if (added.status >= 400 && added.text.toLowerCase().indexOf('already') < 0) {
    assertOk(added);
  }
}

async function registerForms(gateway: ISharePointGateway, listId: string, properties: string): Promise<void> {
  const types = await gateway.get(
    `${gateway.webUrl}/_api/web/lists(guid'${listId}')/contenttypes?$select=Name,StringId`
  );
  assertOk(types);
  const rows = readRows<{ Name?: string; StringId?: string }>(types.text);
  const itemType = rows.filter((row) => row.Name === 'Item')[0] || rows[0];
  if (!itemType || !itemType.StringId) {
    throw new RepeatingRowsError('The list has no content type to attach the form to.');
  }
  await merge(
    gateway,
    `${gateway.webUrl}/_api/web/lists(guid'${listId}')/contenttypes('${itemType.StringId}')`,
    JSON.stringify({
      NewFormClientSideComponentId: FORM_CUSTOMIZER_ID,
      NewFormClientSideComponentProperties: properties,
      EditFormClientSideComponentId: FORM_CUSTOMIZER_ID,
      EditFormClientSideComponentProperties: properties,
      DispFormClientSideComponentId: FORM_CUSTOMIZER_ID,
      DispFormClientSideComponentProperties: properties
    })
  );
}

async function fieldNames(gateway: ISharePointGateway, listId: string): Promise<Set<string>> {
  const result = await gateway.get(
    `${gateway.webUrl}/_api/web/lists(guid'${listId}')/fields?$select=InternalName`
  );
  assertOk(result);
  const names = readRows<{ InternalName?: string }>(result.text).map((field) => (field.InternalName || '').toLowerCase());
  return new Set(names);
}

async function createField(
  gateway: ISharePointGateway,
  listId: string,
  schemaXml: string,
  sendField: (url: string, body: string) => Promise<{ status: number; text: string }>
): Promise<void> {
  const url = `${gateway.webUrl}/_api/web/lists(guid'${listId}')/fields/CreateFieldAsXml`;
  const body = JSON.stringify({ parameters: { SchemaXml: schemaXml, Options: 25 } });
  const result = await sendField(url, body);
  if (result.status >= 400 && result.text.toLowerCase().indexOf('already') < 0 && result.text.toLowerCase().indexOf('duplicate') < 0) {
    assertOk(result);
  }
}

async function merge(gateway: ISharePointGateway, url: string, body: string): Promise<void> {
  const result = await gateway.send('MERGE', url, body, '*');
  assertOk(result);
}

export { fieldHeaders };
