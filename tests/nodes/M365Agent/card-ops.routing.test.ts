import { describe, it, expect, vi, beforeEach } from 'vitest';
import type { BotConnectorBundle } from '../../../shared/botConnector';

vi.mock('../../../shared/auth/router', () => ({
	acquireOutboundToken: vi.fn().mockResolvedValue({ authorizationHeader: 'Bearer routed' }),
}));
vi.mock('../../../shared/botConnector', () => ({
	createConnectorFromBearer: vi.fn().mockImplementation((url: string, header: string) => ({
		client: {
			sendToConversation: vi.fn().mockResolvedValue({ id: 'act-card-send' }),
			updateActivity: vi.fn().mockResolvedValue({ id: 'act-card-update' }),
		},
		axios: {},
		token: header,
		baseURL: url,
	})),
}));

import { execute as cardSendExecute } from '../../../nodes/M365Agent/actions/card/send.operation';
import { execute as cardUpdateExecute } from '../../../nodes/M365Agent/actions/card/update.operation';
import { acquireOutboundToken } from '../../../shared/auth/router';
import { makeExecuteMock } from '../../helpers/executeMock';

const CARD_TEMPLATE = JSON.stringify({
	type: 'AdaptiveCard',
	version: '1.4',
	body: [{ type: 'TextBlock', text: 'hello' }],
});

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

// ── card/send ─────────────────────────────────────────────────────────────────

describe('card/send routing', () => {
	beforeEach(() => (acquireOutboundToken as any).mockClear());

	it('classic authKind → router receives classicBot', async () => {
		const ctx = makeExecuteMock({
			items: ENVELOPE_INPUT,
			nodeParams: (name: string, _i: number, fallback?: unknown) => {
				const map: Record<string, unknown> = {
					authKind: 'classicBot',
					resource: 'card',
					operation: 'send',
					conversationSource: 'envelope',
					cardTemplate: CARD_TEMPLATE,
					bindingData: {},
					options: {},
				};
				return map[name] ?? fallback;
			},
			credentialName: 'm365AgentApi',
			credentials: { appType: 'SingleTenant', clientId: 'c', tenantId: 't', clientSecret: 's' },
		});
		const bundles = new Map<string, BotConnectorBundle>();
		await cardSendExecute.call(
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
					resource: 'card',
					operation: 'send',
					conversationSource: 'envelope',
					cardTemplate: CARD_TEMPLATE,
					bindingData: {},
					options: {},
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
		await cardSendExecute.call(
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
					resource: 'card',
					operation: 'send',
					conversationSource: 'envelope',
					cardTemplate: CARD_TEMPLATE,
					bindingData: {},
					options: {},
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
		await cardSendExecute.call(
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
});

// ── card/update ───────────────────────────────────────────────────────────────

describe('card/update routing', () => {
	beforeEach(() => (acquireOutboundToken as any).mockClear());

	it('classic authKind → router receives classicBot', async () => {
		const ctx = makeExecuteMock({
			items: ENVELOPE_INPUT,
			nodeParams: (name: string, _i: number, fallback?: unknown) => {
				const map: Record<string, unknown> = {
					authKind: 'classicBot',
					resource: 'card',
					operation: 'update',
					conversationSource: 'envelope',
					cardTemplate: CARD_TEMPLATE,
					bindingData: {},
					options: {},
				};
				return map[name] ?? fallback;
			},
			credentialName: 'm365AgentApi',
			credentials: { appType: 'SingleTenant', clientId: 'c', tenantId: 't', clientSecret: 's' },
		});
		const bundles = new Map<string, BotConnectorBundle>();
		await cardUpdateExecute.call(
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
					resource: 'card',
					operation: 'update',
					conversationSource: 'envelope',
					cardTemplate: CARD_TEMPLATE,
					bindingData: {},
					options: {},
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
		await cardUpdateExecute.call(
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
					resource: 'card',
					operation: 'update',
					conversationSource: 'envelope',
					cardTemplate: CARD_TEMPLATE,
					bindingData: {},
					options: {},
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
		await cardUpdateExecute.call(
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
});
