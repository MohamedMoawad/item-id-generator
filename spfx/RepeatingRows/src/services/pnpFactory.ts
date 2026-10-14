import { ISPFXContext, spfi, SPFI, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/batching';
import '@pnp/sp/attachments';
import { IChildRow, IRepeatingRowsConfig } from '../models/repeatingRows';
import { rowPayload } from '../utilities/fieldValues';
import { chunkOperations, BATCH_CHUNK } from '../utilities/batch';

interface ISpfxHost {
  pageContext: {
    web: { absoluteUrl: string };
  };
}

export function createSp(context: ISpfxHost): SPFI {
  return spfi().using(SPFx(context as ISPFXContext));
}

export async function pnpDeleteRows(context: ISpfxHost, config: IRepeatingRowsConfig, ids: number[]): Promise<void> {
  if (ids.length === 0) {
    return;
  }
  const groups = chunkOperations(ids, BATCH_CHUNK);
  for (let index = 0; index < groups.length; index += 1) {
    const sp = createSp(context);
    const [batched, execute] = sp.batched();
    const pending: Promise<void>[] = [];
    groups[index].forEach((id) => {
      pending.push(batched.web.lists.getByTitle(config.childListTitle).items.getById(id).delete());
    });
    await execute();
    await Promise.all(pending.map((item) => item.catch(() => undefined)));
    await execute();
  }
}

export function pnpRowBody(row: IChildRow, config: IRepeatingRowsConfig): { [key: string]: string | number | boolean } {
  return rowPayload(row, config);
}
