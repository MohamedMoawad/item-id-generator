import { MARKER_FIELD_NAME } from '../models/repeatingRows';
import { IChildRow, IPendingAttachment, RepeatingRowsError } from '../models/repeatingRows';
import { ISaveRequest, ISaveResult, IRepeatingStore } from '../repositories/contracts';
import { planChildChanges } from '../utilities/savePlan';
import { validateRows } from '../utilities/validation';

export async function saveRepeatingItem(store: IRepeatingStore, request: ISaveRequest): Promise<ISaveResult> {
  const issues = validateRows(request.rows, store.config, request.attachments);
  if (issues.length > 0) {
    throw new RepeatingRowsError(issues[0].message);
  }
  const parentValues = {
    ...request.parentValues,
    [MARKER_FIELD_NAME]: String(request.rows.length)
  };
  const parentId = request.parentId
    ? request.parentId
    : await store.parent.create(parentValues);
  if (request.parentId) {
    await store.parent.update(request.parentId, parentValues, request.etag);
  }
  const stamped = request.rows.map((row) => ({ ...row, parentId }));
  const plan = planChildChanges(request.originalRows, stamped);
  const created = await store.children.apply(plan);
  const idByKey: { [key: string]: number } = {};
  created.forEach((item) => {
    idByKey[item.key] = item.id;
  });
  stamped.forEach((row) => {
    if (row.id !== undefined) {
      idByKey[row.key] = row.id;
    }
  });
  await writeAttachments(store, request.attachments, idByKey);
  return { parentId, created };
}

async function writeAttachments(
  store: IRepeatingStore,
  attachments: IPendingAttachment[],
  idByKey: { [key: string]: number }
): Promise<void> {
  const groups: { [id: string]: IPendingAttachment[] } = {};
  attachments.forEach((file) => {
    const itemId = idByKey[file.rowKey];
    if (itemId === undefined) {
      return;
    }
    const key = String(itemId);
    groups[key] = groups[key] || [];
    groups[key].push(file);
  });
  const ids = Object.keys(groups);
  for (let index = 0; index < ids.length; index += 1) {
    await store.children.addAttachments(Number(ids[index]), groups[ids[index]]);
  }
}

export function mergeLoadedRows(current: IChildRow[], page: IChildRow[]): IChildRow[] {
  const seen = new Set(current.map((row) => row.id).filter((id) => id !== undefined));
  const next = current.slice();
  page.forEach((row) => {
    if (row.id === undefined || !seen.has(row.id)) {
      next.push(row);
    }
  });
  return next;
}
