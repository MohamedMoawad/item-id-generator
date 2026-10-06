export interface IGenerateItemIdRequest {
  functionUrl: string;
  listName: string;
  condition: string;
  itemId: number;
}

export function buildGenerateItemIdBody(request: IGenerateItemIdRequest): string {
  return JSON.stringify({
    listName: request.listName.trim(),
    condition: request.condition.trim(),
    itemId: request.itemId
  });
}

/**
 * POSTs { listName, condition, itemId } to the Function URL from the property pane.
 * Extra metadata fields can be added to this JSON body when a new ID rule needs them.
 */
export async function requestGeneratedItemId(request: IGenerateItemIdRequest): Promise<string> {
  const functionUrl = request.functionUrl.trim();
  if (!functionUrl || functionUrl.indexOf('<') >= 0) {
    throw new Error('Set the Function URL in the web part property pane. Replace the <function-app> placeholder.');
  }
  if (!request.listName.trim()) {
    throw new Error('Set the list name in the web part property pane.');
  }
  if (!Number.isInteger(request.itemId) || request.itemId <= 0) {
    throw new Error('Enter the SharePoint list item ID (a positive integer).');
  }

  let response: Response;
  try {
    response = await fetch(functionUrl, {
      method: 'POST',
      headers: {
        Accept: 'application/json',
        'Content-Type': 'application/json'
      },
      body: buildGenerateItemIdBody(request)
    });
  } catch (error) {
    const detail = error instanceof Error ? error.message : 'Network request failed.';
    throw new Error(
      `The browser could not reach the Function (${detail}). Check the URL, that the Function is running, and that CORS allows this SharePoint site.`
    );
  }

  const text = await response.text();
  let payload: { id?: unknown; error?: unknown };
  try {
    payload = JSON.parse(text) as { id?: unknown; error?: unknown };
  } catch {
    throw new Error(`The Function returned a non-JSON response (HTTP ${response.status}).`);
  }

  if (!response.ok) {
    const message = typeof payload.error === 'string' && payload.error.trim()
      ? payload.error
      : `The Function returned HTTP ${response.status}.`;
    throw new Error(message);
  }

  if (typeof payload.id !== 'string' || payload.id.trim() === '') {
    throw new Error('The Function response did not include an id string.');
  }

  return payload.id;
}
