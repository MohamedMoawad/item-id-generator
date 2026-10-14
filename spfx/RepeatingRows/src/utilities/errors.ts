import { RepeatingRowsError } from '../models/repeatingRows';

export function toUserMessage(error: unknown): string {
  if (error instanceof RepeatingRowsError) {
    return error.message;
  }
  if (error instanceof Error && error.message) {
    return error.message;
  }
  return 'Something went wrong while saving the rows. Try again.';
}

export function messageFromStatus(status: number, body: string): string {
  if (status === 401 || status === 403) {
    return 'You do not have permission to change these rows.';
  }
  if (status === 404) {
    return 'The list or item was not found.';
  }
  if (status === 412) {
    return 'Someone else changed this item. Reload the form and try again.';
  }
  const parsed = readSharePointMessage(body);
  if (parsed) {
    return parsed;
  }
  return `SharePoint returned status ${status}.`;
}

export function readSharePointMessage(body: string): string {
  if (!body) {
    return '';
  }
  try {
    const parsed = JSON.parse(body) as unknown;
    return digMessage(parsed);
  } catch (error) {
    return error instanceof Error ? '' : '';
  }
}

function digMessage(value: unknown): string {
  if (!value || typeof value !== 'object') {
    return '';
  }
  const record = value as { [key: string]: unknown };
  const direct = record.message;
  if (typeof direct === 'string') {
    return direct;
  }
  if (direct && typeof direct === 'object') {
    const nested = (direct as { value?: unknown }).value;
    if (typeof nested === 'string') {
      return nested;
    }
  }
  const odata = record['odata.error'] || record.error;
  if (odata) {
    return digMessage(odata);
  }
  return '';
}
