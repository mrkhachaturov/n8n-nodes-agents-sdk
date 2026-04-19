import { describe, it, expect, vi, beforeEach } from 'vitest';
import { NodeApiError, NodeOperationError } from 'n8n-workflow';
import type { IExecuteFunctions, IDataObject } from 'n8n-workflow';
import type { Attachment } from '@microsoft/agents-activity';

vi.mock('../../../shared/auth/router', () => ({
	acquireOutboundToken: vi.fn().mockResolvedValue({ authorizationHeader: 'Bearer fake' }),
}));

const fakeClient = { sendToConversation: vi.fn() };
const fakeBundle = { client: fakeClient, axios: {}, token: 'fake', baseURL: 'https://smba/' };
const mockCreate = vi.fn().mockReturnValue(fakeBundle);

vi.mock('../../../shared/botConnector', () => ({
	createConnectorFromBearer: (...args: unknown[]) => mockCreate(...args),
}));

import { dispatchSend } from '../../../shared/card/dispatchSend';
import { makeExecuteContext, makeCredentials } from '../../helpers/makeContext';

const ENVELOPE = {
	conversationReference: {
		serviceUrl: 'https://smba/',
		conversation: { id: '19:c@thread.tacv2' },
		channelId: 'msteams',
	},
};

describe('dispatchSend', () => {
	beforeEach(() => {
		vi.clearAllMocks();
		mockCreate.mockReturnValue(fakeBundle);
	});

	it('sends an attachment via client.sendToConversation and returns { ...item, cardSendResult }', async () => {
		fakeClient.sendToConversation.mockResolvedValue({ id: 'sent-1' });
		const build = vi.fn().mockReturnValue({
			contentType: 'application/vnd.microsoft.card.hero',
			content: { title: 'X' },
		} as Attachment);

		const ctx = makeExecuteContext({
			inputItems: [{ ...ENVELOPE, extra: 'keep' }],
			credentials: makeCredentials(),
			parameters: {
				resource: 'heroCard',
				operation: 'send',
				conversationSource: 'envelope',
				options: {},
				authKind: 'classicBot',
			},
		}) as unknown as IExecuteFunctions;
		const bundles = new Map();
		const credentials = makeCredentials();

		const result = (await dispatchSend({
			ctx,
			itemIndex: 0,
			resourceLabel: 'heroCard',
			authKind: 'classicBot',
			credentials: credentials as never,
			bundles,
			buildAttachment: build,
		})) as IDataObject;

		expect(build).toHaveBeenCalledWith(ctx, 0);
		expect(fakeClient.sendToConversation).toHaveBeenCalledTimes(1);
		expect(result.cardSendResult).toEqual({ id: 'sent-1' });
		expect(result.extra).toBe('keep');
	});

	it('wraps non-NodeOperationError failures from the connector as NodeApiError and includes resourceLabel', async () => {
		fakeClient.sendToConversation.mockRejectedValue(new Error('network down'));
		const build = vi.fn().mockReturnValue({ contentType: 'x', content: {} } as Attachment);

		const ctx = makeExecuteContext({
			inputItems: [ENVELOPE],
			credentials: makeCredentials(),
			parameters: {
				resource: 'heroCard',
				operation: 'send',
				conversationSource: 'envelope',
				options: {},
				authKind: 'classicBot',
			},
		}) as unknown as IExecuteFunctions;

		const promise = dispatchSend({
			ctx,
			itemIndex: 0,
			resourceLabel: 'heroCard',
			authKind: 'classicBot',
			credentials: makeCredentials() as never,
			bundles: new Map(),
			buildAttachment: build,
		});
		await expect(promise).rejects.toBeInstanceOf(NodeApiError);
		await expect(promise).rejects.toThrow(/heroCard send failed/);
	});

	it('lets NodeOperationError from buildAttachment propagate unchanged', async () => {
		const build = vi.fn().mockImplementation(() => {
			throw new NodeOperationError({ name: 'T' } as never, 'bad template');
		});

		const ctx = makeExecuteContext({
			inputItems: [ENVELOPE],
			credentials: makeCredentials(),
			parameters: {
				resource: 'adaptiveCard',
				operation: 'send',
				conversationSource: 'envelope',
				options: {},
				authKind: 'classicBot',
			},
		}) as unknown as IExecuteFunctions;

		await expect(
			dispatchSend({
				ctx,
				itemIndex: 0,
				resourceLabel: 'adaptiveCard',
				authKind: 'classicBot',
				credentials: makeCredentials() as never,
				bundles: new Map(),
				buildAttachment: build,
			}),
		).rejects.toBeInstanceOf(NodeOperationError);
	});

	it('wraps a plain Error from buildAttachment as NodeApiError (card build failed)', async () => {
		const build = vi.fn().mockImplementation(() => {
			throw new Error('unexpected builder crash');
		});

		const ctx = makeExecuteContext({
			inputItems: [ENVELOPE],
			credentials: makeCredentials(),
			parameters: {
				resource: 'heroCard',
				operation: 'send',
				conversationSource: 'envelope',
				options: {},
				authKind: 'classicBot',
			},
		}) as unknown as IExecuteFunctions;

		await expect(
			dispatchSend({
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
		// CardBuildError is thrown by cardActionRow/parseJsonOr on invalid Button /
		// Image JSON. The spec §5.1 requires user-config JSON errors to surface as
		// NodeOperationError (yellow banner), not NodeApiError (red banner).
		const { CardBuildError } = await import('../../../shared/cardBuilders/cardActionRow');
		const build = vi.fn().mockImplementation(() => {
			throw new CardBuildError('Invalid JSON in Value JSON: Unexpected token');
		});

		const ctx = makeExecuteContext({
			inputItems: [ENVELOPE],
			credentials: makeCredentials(),
			parameters: {
				resource: 'heroCard',
				operation: 'send',
				conversationSource: 'envelope',
				options: {},
				authKind: 'classicBot',
			},
		}) as unknown as IExecuteFunctions;

		await expect(
			dispatchSend({
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

	it('adds fallbackText from options to Activity.text', async () => {
		fakeClient.sendToConversation.mockResolvedValue({ id: 'x' });
		const build = vi.fn().mockReturnValue({ contentType: 'x', content: {} } as Attachment);

		const ctx = makeExecuteContext({
			inputItems: [ENVELOPE],
			credentials: makeCredentials(),
			parameters: {
				resource: 'heroCard',
				operation: 'send',
				conversationSource: 'envelope',
				options: { fallbackText: 'fallback' },
				authKind: 'classicBot',
			},
		}) as unknown as IExecuteFunctions;

		await dispatchSend({
			ctx,
			itemIndex: 0,
			resourceLabel: 'heroCard',
			authKind: 'classicBot',
			credentials: makeCredentials() as never,
			bundles: new Map(),
			buildAttachment: build,
		});

		const [, activity] = fakeClient.sendToConversation.mock.calls[0] as [string, { text?: string }];
		expect(activity.text).toBe('fallback');
	});

	it('reuses a cached bundle across calls with the same key', async () => {
		fakeClient.sendToConversation.mockResolvedValue({ id: 'x' });
		const build = vi.fn().mockReturnValue({ contentType: 'x', content: {} } as Attachment);
		const bundles = new Map();

		const ctx = makeExecuteContext({
			inputItems: [ENVELOPE, ENVELOPE],
			credentials: makeCredentials(),
			parameters: {
				resource: 'heroCard',
				operation: 'send',
				conversationSource: 'envelope',
				options: {},
				authKind: 'classicBot',
			},
		}) as unknown as IExecuteFunctions;

		for (const i of [0, 1]) {
			await dispatchSend({
				ctx,
				itemIndex: i,
				resourceLabel: 'heroCard',
				authKind: 'classicBot',
				credentials: makeCredentials() as never,
				bundles,
				buildAttachment: build,
			});
		}
		expect(mockCreate).toHaveBeenCalledTimes(1);
	});
});
