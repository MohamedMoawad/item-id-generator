export function summaryText(value: string | undefined): string {
  const count = Number(value);
  if (!value || !Number.isFinite(count) || count <= 0) {
    return 'No line items';
  }
  return count === 1 ? '1 line item' : `${count} line items`;
}
