import { normalizeGuid } from './contract';

export function buildItemQuery(listGuid: string, numberColumn: string, formulaColumns: string): string {
  const guid = normalizeGuid(listGuid);
  const column = numberColumn.trim();
  const select = ['Id', 'Created', column];
  const expand: string[] = [];
  formulaColumns.split(',').forEach((raw) => {
    const token = raw.trim();
    if (!token) {
      return;
    }
    const name = token.split('.')[0];
    if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) {
      return;
    }
    const lower = name.toLowerCase();
    if (lower === column.toLowerCase() || lower === 'id' || lower === 'created') {
      return;
    }
    if (token.indexOf('.') >= 0) {
      select.push(`${name}/Title`);
      if (expand.indexOf(name) === -1) {
        expand.push(name);
      }
      return;
    }
    select.push(name);
  });
  let query = `_api/web/lists(guid'${guid}')/items?$select=${unique(select).join(',')}&$filter=${column}+eq+null&$orderby=Id+asc&$top=100`;
  if (expand.length > 0) {
    query += `&$expand=${expand.join(',')}`;
  }
  return query;
}

function unique(values: string[]): string[] {
  const seen: { [value: string]: boolean } = {};
  return values.filter((value) => {
    if (seen[value]) {
      return false;
    }
    seen[value] = true;
    return true;
  });
}
