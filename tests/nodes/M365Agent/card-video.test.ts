import { describe, it, expect, vi, beforeEach } from 'vitest';
import { M365Agent } from '../../../nodes/M365Agent/M365Agent.node';
import { makeExecuteContext, makeCredentials } from '../../helpers/makeContext';
import type { IExecuteFunctions } from 'n8n-workflow';

vi.mock('../../../shared/auth/router', () => ({
	acquireOutboundToken: vi.fn().mockResolvedValue({ authorizationHeader: 'Bearer fake' }),
}));

const fakeClient = {
	sendToConversation: vi.fn(),
	updateActivity: vi.fn(),
};
const fakeBundle = { client: fakeClient, axios: {}, token: 'fake', baseURL: 'https://smba/' };
const mockCreate = vi.fn().mockReturnValue(fakeBundle);

vi.mock('../../../shared/botConnector', () => ({
	createConnectorFromBearer: (...args: unknown[]) => mockCreate(...args),
	replyInThread: vi.fn(),
}));

const ENV = {
	conversationReference: {
		serviceUrl: 'https://smba/',
		conversation: { id: '19:c@thread.tacv2' },
		channelId: 'msteams',
	},
	activity: { type: 'message' },
	customField: 'preserved',
};

describe('M365Agent videoCard send', () => {
	beforeEach(() => {
		vi.clearAllMocks();
		mockCreate.mockReturnValue(fakeBundle);
		fakeClient.sendToConversation.mockResolvedValue({ id: 'sent-video' });
	});

	it('sends a video attachment and preserves input', async () => {
		const node = new M365Agent();
		const ctx = makeExecuteContext({
			inputItems: [ENV],
			credentials: makeCredentials(),
			parameters: {
				resource: 'videoCard',
				operation: 'send',
				authKind: 'classicBot',
				conversationSource: 'envelope',
				title: 'Clip',
				media: { mediaItem: [{ url: 'https://m.mp4' }] },
				buttons: {},
				options: { fallbackText: 'Video fallback' },
			},
		});
		const result = await node.execute.call(ctx as unknown as IExecuteFunctions);

		expect(fakeClient.sendToConversation).toHaveBeenCalledTimes(1);
		const [, activity] = fakeClient.sendToConversation.mock.calls[0] as [
			string,
			{ attachments: Array<{ contentType: string }>; text?: string },
		];
		expect(activity.attachments[0].contentType).toBe('application/vnd.microsoft.card.video');
		expect(activity.text).toBe('Video fallback');
		expect(result[0][0].json).toMatchObject({
			customField: 'preserved',
			cardSendResult: { id: 'sent-video' },
		});
	});

	it('rejects a send with no media URLs', async () => {
		const node = new M365Agent();
		const ctx = makeExecuteContext({
			inputItems: [ENV],
			credentials: makeCredentials(),
			parameters: {
				resource: 'videoCard',
				operation: 'send',
				authKind: 'classicBot',
				conversationSource: 'envelope',
				title: 'T',
				media: {},
				buttons: {},
				options: {},
			},
		});
		await expect(node.execute.call(ctx as unknown as IExecuteFunctions)).rejects.toThrow(
			/at least one media URL/,
		);
	});
});
