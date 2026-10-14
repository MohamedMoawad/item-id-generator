export const BATCH_CHUNK = 20;

export function chunkOperations<T>(items: T[], size: number): T[][] {
  const width = size > 0 ? size : items.length || 1;
  const chunks: T[][] = [];
  for (let index = 0; index < items.length; index += width) {
    chunks.push(items.slice(index, index + width));
  }
  return chunks;
}

export interface IBatchRequest {
  method: 'POST' | 'MERGE' | 'DELETE';
  url: string;
  body?: string;
  etag?: string;
}

export function buildBatchPayload(requests: IBatchRequest[], batchId: string): { body: string; contentType: string } {
  const changeId = `changeset_${batchId}`;
  const parts = requests.map((request) => {
    const headers = [
      'Content-Type: application/json;odata=nometadata',
      'Accept: application/json;odata=nometadata',
      'odata-version: '
    ];
    if (request.method === 'MERGE' || request.method === 'DELETE') {
      headers.push(`If-Match: ${request.etag || '*'}`);
    }
    if (request.method === 'MERGE') {
      headers.push('X-HTTP-Method: MERGE');
    }
    const method = request.method === 'MERGE' ? 'POST' : request.method;
    const head = [
      `--${changeId}`,
      'Content-Type: application/http',
      'Content-Transfer-Encoding: binary',
      '',
      `${method} ${request.url} HTTP/1.1`,
      headers.join('\r\n')
    ];
    if (request.body && request.method !== 'DELETE') {
      head.push('', request.body);
    }
    return head.join('\r\n');
  });
  const body = [
    `--batch_${batchId}`,
    `Content-Type: multipart/mixed; boundary=${changeId}`,
    '',
    parts.join('\r\n'),
    `--${changeId}--`,
    `--batch_${batchId}--`,
    ''
  ].join('\r\n');
  return {
    body,
    contentType: `multipart/mixed; boundary=batch_${batchId}`
  };
}

export function readBatchStatuses(responseText: string): number[] {
  const statuses: number[] = [];
  const pattern = /HTTP\/1\.[01] (\d{3})/g;
  let match = pattern.exec(responseText);
  while (match) {
    statuses.push(Number(match[1]));
    match = pattern.exec(responseText);
  }
  return statuses;
}

export function visibleRange(
  count: number,
  scrollTop: number,
  viewport: number,
  rowHeight: number
): { start: number; end: number } {
  const height = rowHeight > 0 ? rowHeight : 1;
  const start = Math.max(0, Math.floor(scrollTop / height) - 2);
  const end = Math.min(count, start + Math.ceil(viewport / height) + 4);
  return { start, end };
}
