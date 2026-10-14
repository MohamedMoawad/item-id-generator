import { IRepeatingRowsConfig, MARKER_FIELD_NAME } from '../models/repeatingRows';
import { parseRepeatingConfig } from '../utilities/configParser';
import { ISharePointGateway, parseJson } from './httpTypes';

export async function loadExistingConfig(
  gateway: ISharePointGateway,
  listId: string,
  fallbackChildList: string
): Promise<IRepeatingRowsConfig | undefined> {
  const result = await gateway.get(
    `${gateway.webUrl}/_api/web/lists(guid'${listId}')/fields/getbyinternalnameortitle('${MARKER_FIELD_NAME}')?$select=ClientSideComponentProperties`
  );
  if (result.status === 404) {
    return undefined;
  }
  if (result.status < 200 || result.status >= 300) {
    return undefined;
  }
  const payload = parseJson<{ ClientSideComponentProperties?: string }>(result.text);
  if (!payload.ClientSideComponentProperties) {
    return undefined;
  }
  try {
    return parseRepeatingConfig(JSON.parse(payload.ClientSideComponentProperties) as unknown, fallbackChildList);
  } catch (error) {
    return error ? undefined : undefined;
  }
}
