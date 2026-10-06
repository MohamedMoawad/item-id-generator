import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IListItemFieldUpdate {
  webAbsoluteUrl: string;
  /** List title from the property pane, for example "Requests". Not the URL slug. */
  listName: string;
  itemId: number;
  /**
   * Internal name of the text column that stores the generated ID.
   * Set this in the property pane. Example: GeneratedItemId.
   * Find it under List settings → column → the Field= value in the URL.
   */
  targetFieldInternalName: string;
  generatedId: string;
}

export interface IListItemMergeRequest {
  url: string;
  body: string;
}

/**
 * Builds the SharePoint REST MERGE request that writes one field.
 * getbytitle() takes the list title with single quotes doubled, not encodeURIComponent.
 */
export function buildListItemMerge(update: IListItemFieldUpdate): IListItemMergeRequest {
  const webAbsoluteUrl = update.webAbsoluteUrl.trim();
  if (!webAbsoluteUrl) {
    throw new Error('The current SharePoint web URL is missing.');
  }

  const listName = update.listName.trim();
  if (!listName) {
    throw new Error('Set the list name in the web part property pane.');
  }

  if (!Number.isInteger(update.itemId) || update.itemId <= 0) {
    throw new Error('Enter the SharePoint list item ID (a positive integer).');
  }

  const fieldInternalName = update.targetFieldInternalName.trim();
  if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(fieldInternalName)) {
    throw new Error('Set a valid target field internal name in the property pane. Example: GeneratedItemId.');
  }

  if (!update.generatedId.trim()) {
    throw new Error('The generated ID is empty, so the list item was not updated.');
  }

  const web = webAbsoluteUrl.replace(/\/+$/, '');
  const encodedTitle = listName.replace(/'/g, "''");
  const url = `${web}/_api/web/lists/getbytitle('${encodedTitle}')/items(${update.itemId})`;
  const body = JSON.stringify({
    [fieldInternalName]: update.generatedId
  });

  return { url, body };
}

/**
 * PATCH-equivalent update of a single list item field via SPHttpClient.
 *
 * SharePoint Online accepts POST plus X-HTTP-Method: MERGE. SPHttpClient adds
 * the request digest for the current web.
 *
 * If the tenant responds that the list item type is required, switch the
 * content type to verbose metadata and set the entity type from:
 *   {web}/_api/web/lists/getbytitle('<listName>')?$select=ListItemEntityTypeFullName
 * Example body:
 *   { __metadata: { type: 'SP.Data.RequestsListItem' }, GeneratedItemId: '<id>' }
 */
export async function updateListItemField(
  spHttpClient: SPHttpClient,
  update: IListItemFieldUpdate
): Promise<void> {
  const merge = buildListItemMerge(update);
  const response: SPHttpClientResponse = await spHttpClient.post(
    merge.url,
    SPHttpClient.configurations.v1,
    {
      headers: {
        Accept: 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE'
      },
      body: merge.body
    }
  );

  if (!response.ok) {
    const detail = (await response.text()).trim().slice(0, 300);
    throw new Error(`Saving the ID to the list item failed (HTTP ${response.status}). ${detail}`);
  }
}
