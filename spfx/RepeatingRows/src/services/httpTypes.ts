import { RepeatingRowsError } from '../models/repeatingRows';
import { messageFromStatus } from '../utilities/errors';

export interface IHttpResult {
  status: number;
  text: string;
}

export interface ISharePointGateway {
  webUrl: string;
  get(url: string): Promise<IHttpResult>;
  send(method: 'POST' | 'MERGE' | 'DELETE', url: string, body?: string, etag?: string): Promise<IHttpResult>;
  batch(body: string, contentType: string): Promise<IHttpResult>;
  postBinary(url: string, body: ArrayBuffer): Promise<IHttpResult>;
}

export function assertOk(result: IHttpResult): void {
  if (result.status >= 200 && result.status < 300) {
    return;
  }
  throw new RepeatingRowsError(messageFromStatus(result.status, result.text), result.status);
}

export function parseJson<T>(text: string): T {
  if (!text) {
    return {} as T;
  }
  return JSON.parse(text) as T;
}

export function odataTitle(title: string): string {
  return title.replace(/'/g, "''");
}

export function readId(text: string): number {
  const payload = parseJson<{ Id?: number; ID?: number; d?: { Id?: number } }>(text);
  const id = payload.Id || payload.ID || (payload.d && payload.d.Id);
  const numeric = Number(id);
  if (!Number.isFinite(numeric)) {
    throw new RepeatingRowsError('SharePoint did not return the new item id.');
  }
  return numeric;
}

export function readRows<T>(text: string): T[] {
  const payload = parseJson<{ value?: T[]; d?: { results?: T[] } }>(text);
  if (payload.value) {
    return payload.value;
  }
  return (payload.d && payload.d.results) || [];
}
