import { describe, it, expect, vi, beforeEach } from 'vitest';
import type { BotConnectorBundle } from '../../../shared/botConnector';

vi.mock('../../../shared/auth/router', () => ({
	acquireOutboundToken: vi.fn().mockResolvedValue({ authorizationHeader: 'Bearer routed' }),
}));
vi.mock('../../../shared/botConnector', () => ({
	createConnectorFromBearer: vi.fn().mockImplementation((url: string, header: string) => ({
		client: { sendToConversation: vi.fn().mockResolvedValue({ id: 'act-1' }) },
		axios: { post: vi.fn() },
		token: header,
		baseURL: url,
	})),
}));

import { execute as sendExecute } from '../../../nodes/M365Agent/actions/message/send.operation';
import { acquireOutboundToken } from '../../../shared/auth/router';
import { makeExecuteMock } from '../../helpers/executeMock';
import { makeBundleKey } from '../../../nodes/M365Agent/actions/router';

const ENVELOPE_INPUT = [
	{
		json: {
			conversationReference: {
				serviceUrl: 'https://smba/',
				conversation: { id: '19:c1', conversationType: 'personal' },
				channelId: 'msteams',
			},
			activity: { type: 'message', text: 'hi' },
		},
	},
];

describe('message/send routing', () => {
	beforeEach(() => (acquireOutboundToken as any).mockClear());

	it('classic authKind → router receives classicBot', async () => {
		const ctx = makeExecuteMock({
			items: ENVELOPE_INPUT,
			nodeParams: (name: string, _i: number, fallback?: unknown) => {
				const map: Record<string, unknown> = {
					authKind: 'classicBot',
					resource: 'message',
					operation: 'send',
					conversationSource: 'fromEnvelope',
					messageText: 'hi',
				};
				return map[name] ?? fallback;
			},
			credentialName: 'm365AgentApi',
			credentials: { appType: 'SingleTenant', clientId: 'c', tenantId: 't', clientSecret: 's' },
		});
		const bundles = new Map<string, BotConnectorBundle>();
		await sendExecute.call(
			ctx,
			0,
			'classicBot' as any,
			{ appType: 'SingleTenant', clientId: 'c' } as any,
			bundles,
		);
		expect(acquireOutboundToken).toHaveBeenCalledWith(
			expect.objectContaining({ authKind: 'classicBot' }),
		);
	});

	it('agent365 + autonomous → router receives downstreamApi=MessagingBotApi', async () => {
		const ctx = makeExecuteMock({
			items: ENVELOPE_INPUT,
			nodeParams: (name: string, _i: number, fallback?: unknown) => {
				const map: Record<string, unknown> = {
					authKind: 'agent365',
					identityMode: 'autonomous',
					resource: 'message',
					operation: 'send',
					conversationSource: 'fromEnvelope',
					messageText: 'hi',
				};
				return map[name] ?? fallback;
			},
			credentialName: 'm365Agent365Api',
			credentials: {
				tenantId: 't',
				blueprintAppId: 'bp',
				transport: 'sidecar',
				sidecarUrl: 'http://s:5000',
				outboundDownstreamApi: 'MessagingBotApi',
				validateVia: 'sameAsOutbound',
			},
		});
		await sendExecute.call(
			ctx,
			0,
			'agent365' as any,
			{
				tenantId: 't',
				blueprintAppId: 'bp',
				transport: 'sidecar',
				sidecarUrl: 'http://s:5000',
				outboundDownstreamApi: 'MessagingBotApi',
				validateVia: 'sameAsOutbound',
			} as any,
			new Map(),
		);
		expect(acquireOutboundToken).toHaveBeenCalledWith(
			expect.objectContaining({
				authKind: 'agent365',
				identityMode: 'autonomous',
				downstreamApi: 'MessagingBotApi',
			}),
		);
	});

	it('agent365 + agentUser by UPN → forwards agentUsername', async () => {
		const ctx = makeExecuteMock({
			items: ENVELOPE_INPUT,
			nodeParams: (name: string, _i: number, fallback?: unknown) => {
				const map: Record<string, unknown> = {
					authKind: 'agent365',
					identityMode: 'agentUser',
					userSelectorMode: 'byUpn',
					agentUsername: 'a@b.com',
					resource: 'message',
					operation: 'send',
					conversationSource: 'fromEnvelope',
					messageText: 'hi',
				};
				return map[name] ?? fallback;
			},
			credentialName: 'm365Agent365Api',
			credentials: {
				tenantId: 't',
				blueprintAppId: 'bp',
				transport: 'sidecar',
				sidecarUrl: 'http://s:5000',
				outboundDownstreamApi: 'MessagingBotApi',
				validateVia: 'sameAsOutbound',
			},
		});
		await sendExecute.call(
			ctx,
			0,
			'agent365' as any,
			{
				tenantId: 't',
				blueprintAppId: 'bp',
				transport: 'sidecar',
				sidecarUrl: 'http://s:5000',
				outboundDownstreamApi: 'MessagingBotApi',
				validateVia: 'sameAsOutbound',
			} as any,
			new Map(),
		);
		expect(acquireOutboundToken).toHaveBeenCalledWith(
			expect.objectContaining({
				identityMode: 'agentUser',
				agentUsername: 'a@b.com',
			}),
		);
	});

	it('uses makeBundleKey to cache — different identityMode produces different key', () => {
		expect(makeBundleKey('https://x/', 'agent365', 'autonomous')).not.toEqual(
			makeBundleKey('https://x/', 'agent365', 'agentUser', 'u@x.com'),
		);
	});
});
