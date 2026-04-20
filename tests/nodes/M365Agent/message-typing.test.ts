import { describe, it, expect, vi, beforeEach } from 'vitest';
import type { BotConnectorBundle } from '../../../shared/botConnector';

// Shared client mock so tests can inspect the call args after execute().
const fakeClient = {
	sendToConversation: vi.fn(),
};

vi.mock('../../../shared/auth/router', () => ({
	acquireOutboundToken: vi.fn().mockResolvedValue({ authorizationHeader: 'Bearer routed' }),
}));
vi.mock('../../../shared/botConnector', () => ({
	createConnectorFromBearer: vi.fn().mockImplementation((url: string, header: string) => ({
		client: fakeClient,
		axios: { post: vi.fn() },
		token: header,
		baseURL: url,
	})),
}));

import { execute as typingExecute } from '../../../nodes/M365Agent/actions/message/typing.operation';
import { acquireOutboundToken } from '../../../shared/auth/router';
import { makeExecuteMock } from '../../helpers/executeMock';

const ENVELOPE_INPUT = [
	{
		json: {
			conversationReference: {
				serviceUrl: 'https://smba/',
				conversation: { id: '19:c1', conversationType: 'personal' },
				channelId: 'msteams',
				activityId: 'act-1',
			},
			activity: { type: 'message', text: 'hi' },
		},
	},
];

describe('message/typing operation', () => {
	beforeEach(() => {
		(acquireOutboundToken as any).mockClear();
		fakeClient.sendToConversation.mockReset();
	});

	it('builds Activity with type=typing and no text/attachments', async () => {
		fakeClient.sendToConversation.mockResolvedValue({ id: 'act-typing-1' });
		const ctx = makeExecuteMock({
			items: ENVELOPE_INPUT,
			nodeParams: (name: string, _i: number, fallback?: unknown) => {
				const map: Record<string, unknown> = {
					authKind: 'classicBot',
					resource: 'message',
					operation: 'typing',
					conversationSource: 'envelope',
				};
				return map[name] ?? fallback;
			},
			credentialName: 'm365AgentApi',
			credentials: { appType: 'SingleTenant', clientId: 'c', tenantId: 't', clientSecret: 's' },
		});
		const bundles = new Map<string, BotConnectorBundle>();
		await typingExecute.call(
			ctx,
			0,
			'classicBot' as any,
			{ appType: 'SingleTenant', clientId: 'c' } as any,
			bundles,
		);

		expect(fakeClient.sendToConversation).toHaveBeenCalledTimes(1);
		const [conversationId, activity] = fakeClient.sendToConversation.mock.calls[0];
		expect(conversationId).toBe('19:c1');
		expect(activity).toEqual({ type: 'typing' });
		// Typing activities must NOT carry text or attachments.
		expect((activity as { text?: unknown }).text).toBeUndefined();
		expect((activity as { attachments?: unknown }).attachments).toBeUndefined();
	});

	it('throws NodeOperationError when conversation reference missing on envelope source', async () => {
		const ctx = makeExecuteMock({
			items: [{ json: {} }],
			nodeParams: (name: string, _i: number, fallback?: unknown) => {
				const map: Record<string, unknown> = {
					authKind: 'classicBot',
					resource: 'message',
					operation: 'typing',
					conversationSource: 'envelope',
				};
				return map[name] ?? fallback;
			},
			credentialName: 'm365AgentApi',
			credentials: { appType: 'SingleTenant', clientId: 'c', tenantId: 't', clientSecret: 's' },
		});
		await expect(
			typingExecute.call(
				ctx,
				0,
				'classicBot' as any,
				{ appType: 'SingleTenant', clientId: 'c' } as any,
				new Map(),
			),
		).rejects.toThrow(
			/Typing requires conversation reference\. Use From Envelope \(default\) on a reply flow, or Specify Manually for proactive\./,
		);
	});

	it('classic authKind → router receives classicBot', async () => {
		fakeClient.sendToConversation.mockResolvedValue({ id: 'act-typing-2' });
		const ctx = makeExecuteMock({
			items: ENVELOPE_INPUT,
			nodeParams: (name: string, _i: number, fallback?: unknown) => {
				const map: Record<string, unknown> = {
					authKind: 'classicBot',
					resource: 'message',
					operation: 'typing',
					conversationSource: 'envelope',
				};
				return map[name] ?? fallback;
			},
			credentialName: 'm365AgentApi',
			credentials: { appType: 'SingleTenant', clientId: 'c', tenantId: 't', clientSecret: 's' },
		});
		await typingExecute.call(
			ctx,
			0,
			'classicBot' as any,
			{ appType: 'SingleTenant', clientId: 'c' } as any,
			new Map(),
		);
		expect(acquireOutboundToken).toHaveBeenCalledWith(
			expect.objectContaining({ authKind: 'classicBot' }),
		);
	});
});
