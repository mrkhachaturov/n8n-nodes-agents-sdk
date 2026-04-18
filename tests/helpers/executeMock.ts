import type { IExecuteFunctions, INode, INodeExecutionData } from 'n8n-workflow';

export interface ExecuteMockInit {
  items?: INodeExecutionData[];
  nodeParams?: Record<string, unknown> | ((name: string, itemIndex: number, fallback?: unknown) => unknown);
  credentials?: Record<string, unknown>;
  credentialName?: string;
}

export function makeExecuteMock(init: ExecuteMockInit): IExecuteFunctions {
  const items = init.items ?? [{ json: {} }];
  const getParam = typeof init.nodeParams === 'function'
    ? init.nodeParams
    : (name: string, _i: number, fallback?: unknown) =>
        (init.nodeParams as Record<string, unknown> | undefined)?.[name] ?? fallback;
  return {
    getInputData: () => items,
    getNodeParameter: getParam as any,
    getCredentials: (async (name: string) => {
      if (init.credentialName && name !== init.credentialName) {
        throw new Error(`Unexpected getCredentials("${name}"), expected "${init.credentialName}"`);
      }
      return init.credentials ?? {};
    }) as any,
    getNode: () => ({ id: 'test-node', name: 'test', type: 'test', typeVersion: 1, position: [0, 0], parameters: {} } as INode),
    helpers: {
      constructExecutionMetaData: (d: any, _m: any) => d,
      returnJsonArray: (d: unknown[]) => d.map(json => ({ json })),
    } as any,
    sendResponse: () => undefined,
  } as unknown as IExecuteFunctions;
}
