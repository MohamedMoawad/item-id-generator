import { IPersonOption } from '../models/repeatingRows';
import { assertOk, ISharePointGateway, parseJson, readId } from './httpTypes';

export async function searchPeople(gateway: ISharePointGateway, text: string): Promise<IPersonOption[]> {
  const query = text.trim();
  if (query.length < 2) {
    return [];
  }
  const result = await gateway.send(
    'POST',
    `${gateway.webUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`,
    JSON.stringify({
      queryParams: {
        AllowEmailAddresses: true,
        AllowMultipleEntities: false,
        AllUrlZones: false,
        MaximumEntitySuggestions: 8,
        PrincipalSource: 15,
        PrincipalType: 1,
        QueryString: query
      }
    })
  );
  assertOk(result);
  const payload = parseJson<{ ClientPeoplePickerSearchUser?: string; d?: { ClientPeoplePickerSearchUser?: string } }>(result.text);
  const raw = payload.ClientPeoplePickerSearchUser || (payload.d && payload.d.ClientPeoplePickerSearchUser) || '[]';
  const people = JSON.parse(raw) as { DisplayText?: string; Key?: string; EntityData?: { SPUserID?: string } }[];
  return people.map((person) => ({
    id: (person.EntityData && person.EntityData.SPUserID) || '',
    login: person.Key || '',
    text: person.DisplayText || person.Key || ''
  }));
}

export async function ensurePerson(gateway: ISharePointGateway, login: string): Promise<number> {
  const result = await gateway.send(
    'POST',
    `${gateway.webUrl}/_api/web/ensureuser`,
    JSON.stringify({ logonName: login })
  );
  assertOk(result);
  return readId(result.text);
}
