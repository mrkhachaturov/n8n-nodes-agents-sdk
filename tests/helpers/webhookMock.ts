import type { IncomingHttpHeaders } from 'http';
import type { IWebhookFunctions, INode } from 'n8n-workflow';

export interface WebhookMockInit {
  body?: unknown;
  headers?: Record<string, string>;
  nodeParams?: Record<string, unknown>;
  credentials?: Record<string, unknown>;
  credentialName?: string;
}

export function makeWebhookMock(init: WebhookMockInit): IWebhookFunctions {
  const req: any = {
    body: init.body ?? {},
    headers: (init.headers ?? {}) as IncomingHttpHeaders,
    method: 'POST',
  };
  return {
    getRequestObject: () => req,
    getResponseObject: () => ({ status: () => ({ json: () => undefined }) }),
    getHeaderData: () => req.headers,
    getBodyData: () => req.body,
    getNodeParameter: ((name: string, fallback?: unknown) =>
      init.nodeParams?.[name] ?? fallback) as any,
    getCredentials: (async (name: string) => {
      if (init.credentialName && name !== init.credentialName) {
        throw new Error(`Unexpected getCredentials("${name}"), expected "${init.credentialName}"`);
      }
      return init.credentials ?? {};
    }) as any,
    getNode: () => ({ id: 'test-node', name: 'test', type: 'test', typeVersion: 1, position: [0, 0], parameters: {} } as INode),
    helpers: { returnJsonArray: (d: unknown[]) => d.map(json => ({ json })) } as any,
    getWebhookName: () => 'default',
  } as unknown as IWebhookFunctions;
}
