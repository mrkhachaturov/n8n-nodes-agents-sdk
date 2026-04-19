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

const HERO_ENVELOPE = {
	conversationReference: {
		serviceUrl: 'https://smba/',
		conversation: { id: '19:c@thread.tacv2' },
		channelId: 'msteams',
	},
	activity: { type: 'message' },
	customField: 'preserved',
};

describe('M365Agent heroCard send', () => {
	beforeEach(() => {
		vi.clearAllMocks();
		mockCreate.mockReturnValue(fakeBundle);
		fakeClient.sendToConversation.mockResolvedValue({ id: 'sent-hero' });
	});

	it('sends a hero attachment with title/subtitle/buttons and preserves input', async () => {
		const node = new M365Agent();
		const ctx = makeExecuteContext({
			inputItems: [HERO_ENVELOPE],
			credentials: makeCredentials(),
			parameters: {
				resource: 'heroCard',
				operation: 'send',
				authKind: 'classicBot',
				conversationSource: 'envelope',
				title: 'Order #42',
				subtitle: 'Ready for pickup',
				text: 'Customer approved',
				images: {},
				buttons: {
					button: [{ type: 'openUrl', title: 'View', value: 'https://view', options: {} }],
				},
				options: { fallbackText: 'Hero fallback' },
			},
		});
		const result = await node.execute.call(ctx as unknown as IExecuteFunctions);

		expect(fakeClient.sendToConversation).toHaveBeenCalledTimes(1);
		const [, activity] = fakeClient.sendToConversation.mock.calls[0] as [
			string,
			{
				attachments: Array<{ contentType: string; content: { title: string; buttons: unknown[] } }>;
				text?: string;
			},
		];
		expect(activity.attachments[0].contentType).toBe('application/vnd.microsoft.card.hero');
		expect(activity.attachments[0].content.title).toBe('Order #42');
		expect(activity.attachments[0].content.buttons).toHaveLength(1);
		expect(activity.text).toBe('Hero fallback');
		expect(result[0][0].json).toMatchObject({
			customField: 'preserved',
			cardSendResult: { id: 'sent-hero' },
		});
	});
});
