import { describe, it, expect, vi, beforeEach } from 'vitest';
import { NodeApiError, NodeOperationError } from 'n8n-workflow';
import type { IExecuteFunctions, IDataObject } from 'n8n-workflow';
import type { Attachment } from '@microsoft/agents-activity';

vi.mock('../../../shared/auth/router', () => ({
  acquireOutboundToken: vi.fn().mockResolvedValue({ authorizationHeader: 'Bearer fake' }),
}));

const fakeClient = { updateActivity: vi.fn() };
const fakeBundle = { client: fakeClient, axios: {}, token: 'fake', baseURL: 'https://smba/' };
const mockCreate = vi.fn().mockReturnValue(fakeBundle);

vi.mock('../../../shared/botConnector', () => ({
  createConnectorFromBearer: (...args: unknown[]) => mockCreate(...args),
}));

import { dispatchUpdate } from '../../../shared/card/dispatchUpdate';
import { makeExecuteContext, makeCredentials } from '../../helpers/makeContext';

const ENVELOPE_WITH_ACT = {
  conversationReference: {
    serviceUrl: 'https://smba/',
    conversation: { id: '19:c@thread.tacv2' },
    activityId: 'act-1',
    channelId: 'msteams',
  },
};
const ENVELOPE_NO_ACT = {
  conversationReference: {
    serviceUrl: 'https://smba/',
    conversation: { id: '19:c@thread.tacv2' },
    channelId: 'msteams',
  },
};

describe('dispatchUpdate', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    mockCreate.mockReturnValue(fakeBundle);
  });

  it('calls client.updateActivity with (conversationId, activityId, activity)', async () => {
    fakeClient.updateActivity.mockResolvedValue({ id: 'updated-1' });
    const build = vi.fn().mockReturnValue({ contentType: 'x', content: {} } as Attachment);

    const ctx = makeExecuteContext({
      inputItems: [ENVELOPE_WITH_ACT],
      credentials: makeCredentials(),
      parameters: {
        resource: 'heroCard',
        operation: 'update',
        conversationSource: 'envelope',
        options: {},
        authKind: 'classicBot',
      },
    }) as unknown as IExecuteFunctions;

    const result = (await dispatchUpdate({
      ctx,
      itemIndex: 0,
      resourceLabel: 'heroCard',
      authKind: 'classicBot',
      credentials: makeCredentials() as never,
      bundles: new Map(),
      buildAttachment: build,
    })) as IDataObject;

    const [convId, actId] = fakeClient.updateActivity.mock.calls[0];
    expect(convId).toBe('19:c@thread.tacv2');
    expect(actId).toBe('act-1');
    expect(result.cardUpdateResult).toEqual({ id: 'updated-1' });
  });

  it('throws NodeOperationError when activityId is missing', async () => {
    const build = vi.fn();
    const ctx = makeExecuteContext({
      inputItems: [ENVELOPE_NO_ACT],
      credentials: makeCredentials(),
      parameters: {
        resource: 'heroCard',
        operation: 'update',
        conversationSource: 'envelope',
        options: {},
        authKind: 'classicBot',
      },
    }) as unknown as IExecuteFunctions;

    await expect(
      dispatchUpdate({
        ctx,
        itemIndex: 0,
        resourceLabel: 'heroCard',
        authKind: 'classicBot',
        credentials: makeCredentials() as never,
        bundles: new Map(),
        buildAttachment: build,
      }),
    ).rejects.toThrow(/Update requires activityId/);
    expect(build).not.toHaveBeenCalled();
  });

  it('wraps a plain Error from buildAttachment as NodeApiError (card build failed)', async () => {
    const build = vi.fn().mockImplementation(() => {
      throw new Error('builder crashed');
    });
    const ctx = makeExecuteContext({
      inputItems: [ENVELOPE_WITH_ACT],
      credentials: makeCredentials(),
      parameters: {
        resource: 'heroCard',
        operation: 'update',
        conversationSource: 'envelope',
        options: {},
        authKind: 'classicBot',
      },
    }) as unknown as IExecuteFunctions;

    await expect(
      dispatchUpdate({
        ctx,
        itemIndex: 0,
        resourceLabel: 'heroCard',
        authKind: 'classicBot',
        credentials: makeCredentials() as never,
        bundles: new Map(),
        buildAttachment: build,
      }),
    ).rejects.toBeInstanceOf(NodeApiError);
  });

  it('reclassifies CardBuildError from a shared builder as NodeOperationError', async () => {
    const { CardBuildError } = await import('../../../shared/cardBuilders/cardActionRow');
    const build = vi.fn().mockImplementation(() => {
      throw new CardBuildError('Invalid JSON in Channel Data: Unexpected token');
    });
    const ctx = makeExecuteContext({
      inputItems: [ENVELOPE_WITH_ACT],
      credentials: makeCredentials(),
      parameters: {
        resource: 'heroCard',
        operation: 'update',
        conversationSource: 'envelope',
        options: {},
        authKind: 'classicBot',
      },
    }) as unknown as IExecuteFunctions;

    await expect(
      dispatchUpdate({
        ctx,
        itemIndex: 0,
        resourceLabel: 'heroCard',
        authKind: 'classicBot',
        credentials: makeCredentials() as never,
        bundles: new Map(),
        buildAttachment: build,
      }),
    ).rejects.toBeInstanceOf(NodeOperationError);
  });

  it('wraps connector failures as NodeApiError and includes resourceLabel', async () => {
    fakeClient.updateActivity.mockRejectedValue(new Error('nope'));
    const build = vi.fn().mockReturnValue({ contentType: 'x', content: {} } as Attachment);
    const ctx = makeExecuteContext({
      inputItems: [ENVELOPE_WITH_ACT],
      credentials: makeCredentials(),
      parameters: {
        resource: 'heroCard',
        operation: 'update',
        conversationSource: 'envelope',
        options: {},
        authKind: 'classicBot',
      },
    }) as unknown as IExecuteFunctions;

    const promise = dispatchUpdate({
      ctx,
      itemIndex: 0,
      resourceLabel: 'heroCard',
      authKind: 'classicBot',
      credentials: makeCredentials() as never,
      bundles: new Map(),
      buildAttachment: build,
    });
    await expect(promise).rejects.toBeInstanceOf(NodeApiError);
    await expect(promise).rejects.toThrow(/heroCard update failed/);
  });
});
