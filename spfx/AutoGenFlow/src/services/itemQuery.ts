export function buildItemQuery(numberColumn: string, formulaColumns: string): string {
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
  let query = `$select=${unique(select).join(',')}`;
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
