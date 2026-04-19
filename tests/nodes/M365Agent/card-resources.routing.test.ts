import { describe, it, expect, vi, beforeEach } from 'vitest';
import { M365Agent } from '../../../nodes/M365Agent/M365Agent.node';
import { makeExecuteContext, makeCredentials } from '../../helpers/makeContext';
import type { IExecuteFunctions } from 'n8n-workflow';

vi.mock('../../../shared/auth/router', () => ({
	acquireOutboundToken: vi.fn().mockResolvedValue({ authorizationHeader: 'Bearer fake' }),
}));

const fakeClient = {
	sendToConversation: vi.fn().mockResolvedValue({ id: 'ok' }),
	updateActivity: vi.fn().mockResolvedValue({ id: 'ok' }),
};
const fakeBundle = { client: fakeClient, axios: {}, token: 'fake', baseURL: 'https://smba/' };
const mockCreate = vi.fn().mockReturnValue(fakeBundle);

vi.mock('../../../shared/botConnector', () => ({
	createConnectorFromBearer: (...args: unknown[]) => mockCreate(...args),
	replyInThread: vi.fn(),
}));

const ENVELOPE = {
	conversationReference: {
		serviceUrl: 'https://smba/',
		conversation: { id: '19:c@thread.tacv2' },
		channelId: 'msteams',
	},
};

/** Per-resource minimal field bag that satisfies each resource's required inputs. */
const MIN_FIELDS: Record<string, Record<string, unknown>> = {
	adaptiveCard: {
		cardTemplate: '{"type":"AdaptiveCard","version":"1.4","body":[]}',
		bindingData: {},
		options: {},
	},
	animationCard: {
		title: 'T',
		media: { mediaItem: [{ url: 'https://m.gif' }] },
		buttons: {},
		options: {},
	},
	audioCard: {
		title: 'T',
		media: { mediaItem: [{ url: 'https://m.mp3' }] },
		buttons: {},
		options: {},
	},
	heroCard: { title: 'T', images: {}, buttons: {}, options: {} },
	o365ConnectorCard: { cardContent: '{"title":"x"}', options: {} },
	rawAttachment: { contentType: 'x', content: '={}', options: {} },
	receiptCard: { cardContent: '{"title":"R"}', options: {} },
	signInCard: { title: 'Sign in', url: 'https://a', text: '' },
	thumbnailCard: { title: 'T', images: {}, buttons: {}, options: {} },
	videoCard: {
		title: 'T',
		media: { mediaItem: [{ url: 'https://m.mp4' }] },
		buttons: {},
		options: {},
	},
};

describe('card resource routing', () => {
	beforeEach(() => {
		vi.clearAllMocks();
		mockCreate.mockReturnValue(fakeBundle);
		fakeClient.sendToConversation.mockResolvedValue({ id: 'ok' });
		fakeClient.updateActivity.mockResolvedValue({ id: 'ok' });
	});

	for (const resource of Object.keys(MIN_FIELDS)) {
		it(`routes ${resource} send to client.sendToConversation`, async () => {
			const node = new M365Agent();
			const ctx = makeExecuteContext({
				inputItems: [ENVELOPE],
				credentials: makeCredentials(),
				parameters: {
					resource,
					operation: 'send',
					authKind: 'classicBot',
					conversationSource: 'envelope',
					...MIN_FIELDS[resource],
				},
			});
			await node.execute.call(ctx as unknown as IExecuteFunctions);
			expect(fakeClient.sendToConversation).toHaveBeenCalledTimes(1);
		});

		it(`routes ${resource} update to client.updateActivity`, async () => {
			const envWithAct = {
				...ENVELOPE,
				conversationReference: { ...ENVELOPE.conversationReference, activityId: 'act-1' },
			};
			const node = new M365Agent();
			const ctx = makeExecuteContext({
				inputItems: [envWithAct],
				credentials: makeCredentials(),
				parameters: {
					resource,
					operation: 'update',
					authKind: 'classicBot',
					conversationSource: 'envelope',
					...MIN_FIELDS[resource],
				},
			});
			await node.execute.call(ctx as unknown as IExecuteFunctions);
			expect(fakeClient.updateActivity).toHaveBeenCalledTimes(1);
		});

		it(`throws NodeOperationError for unknown ${resource} operation`, async () => {
			const node = new M365Agent();
			const ctx = makeExecuteContext({
				inputItems: [ENVELOPE],
				credentials: makeCredentials(),
				parameters: {
					resource,
					operation: 'bogus',
					authKind: 'classicBot',
					conversationSource: 'envelope',
					...MIN_FIELDS[resource],
				},
			});
			await expect(node.execute.call(ctx as unknown as IExecuteFunctions)).rejects.toThrow(
				new RegExp(`Unknown ${resource} operation`),
			);
		});
	}

	// ---- Router error-model coverage (spec §5.1 — 6a, 6c) -----------------
	//
	// 6a. invokeResponse early-return guard — unknown invokeResponse operation.
	//     This lives outside the per-item loop, so it needs its own test.
	it('6a: invokeResponse with unknown operation throws NodeOperationError', async () => {
		const node = new M365Agent();
		const ctx = makeExecuteContext({
			inputItems: [ENVELOPE],
			credentials: makeCredentials(),
			parameters: {
				resource: 'invokeResponse',
				operation: 'bogus',
				authKind: 'classicBot',
				conversationSource: 'envelope',
			},
		});
		await expect(node.execute.call(ctx as unknown as IExecuteFunctions)).rejects.toThrow(
			/Unknown invokeResponse operation: bogus/,
		);
	});

	// 6c. Per-item default branch — an unexpected resource string that is not in
	//     the per-item switch statement. Forces the router's default clause to fire.
	//     This is the same path that the 0.3.x regression test (Task 21) pins.
	it('6c: an unknown resource hits the router default with the preserved invariant message', async () => {
		const node = new M365Agent();
		const ctx = makeExecuteContext({
			inputItems: [ENVELOPE],
			credentials: makeCredentials(),
			parameters: {
				resource: 'nonexistentCard',
				operation: 'send',
				authKind: 'classicBot',
				conversationSource: 'envelope',
			},
		});
		const promise = node.execute.call(ctx as unknown as IExecuteFunctions);
		await expect(promise).rejects.toThrow(/Unexpected resource "nonexistentCard"/);
		await expect(promise).rejects.toThrow(
			/Invoke Response must be the only operation in this node execution/,
		);
	});
});
