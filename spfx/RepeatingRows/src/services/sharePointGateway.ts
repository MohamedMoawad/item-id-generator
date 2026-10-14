import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IHttpResult, ISharePointGateway } from './httpTypes';

const jsonHeaders: { [key: string]: string } = {
  Accept: 'application/json;odata=nometadata',
  'Content-Type': 'application/json;odata=nometadata',
  'odata-version': ''
};

export class SpHttpGateway implements ISharePointGateway {
  public constructor(private readonly client: SPHttpClient, public readonly webUrl: string) {}

  public async get(url: string): Promise<IHttpResult> {
    const response = await this.client.get(url, SPHttpClient.configurations.v1, {
      headers: { Accept: 'application/json;odata=nometadata', 'odata-version': '' }
    });
    return readResult(response);
  }

  public async send(method: 'POST' | 'MERGE' | 'DELETE', url: string, body?: string, etag?: string): Promise<IHttpResult> {
    const headers: { [key: string]: string } = { ...jsonHeaders };
    if (method === 'MERGE' || method === 'DELETE') {
      headers['IF-MATCH'] = etag || '*';
      headers['X-HTTP-Method'] = method === 'MERGE' ? 'MERGE' : 'DELETE';
    }
    const response = await this.client.post(url, SPHttpClient.configurations.v1, { headers, body: body || '{}' });
    return readResult(response);
  }

  public async batch(body: string, contentType: string): Promise<IHttpResult> {
    const response = await this.client.post(`${this.webUrl}/_api/$batch`, SPHttpClient.configurations.v1, {
      headers: {
        Accept: 'application/json',
        'Content-Type': contentType,
        'odata-version': ''
      },
      body
    });
    return readResult(response);
  }

  public async postBinary(url: string, body: ArrayBuffer): Promise<IHttpResult> {
    const response = await this.client.post(url, SPHttpClient.configurations.v1, {
      headers: {
        Accept: 'application/json;odata=nometadata',
        'Content-Type': 'application/octet-stream',
        'odata-version': ''
      },
      body
    });
    return readResult(response);
  }
}

async function readResult(response: SPHttpClientResponse): Promise<IHttpResult> {
  return { status: response.status, text: await response.text() };
}
