import { describe, it, expect, vi, beforeEach } from 'vitest';
import type { BotConnectorBundle } from '../../../shared/botConnector';

vi.mock('../../../shared/auth/router', () => ({
	acquireOutboundToken: vi.fn().mockResolvedValue({ authorizationHeader: 'Bearer routed' }),
}));

const sendToConversation = vi.fn().mockResolvedValue({ id: 'act-cardstate-send' });

vi.mock('../../../shared/botConnector', () => ({
	createConnectorFromBearer: vi.fn().mockImplementation((url: string, header: string) => ({
		client: {
			sendToConversation,
			updateActivity: vi.fn().mockResolvedValue({ id: 'act-cardstate-update' }),
		},
		axios: {},
		token: header,
		baseURL: url,
	})),
}));

import { execute as cardStateSelectExecute } from '../../../nodes/M365Agent/actions/card-state/select.operation';
import { makeExecuteMock } from '../../helpers/executeMock';

type Variant = { key: string; cardJson: string; bindingData?: string };

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

function makeCtxFor(
	stateKey: string,
	variants: Variant[],
	overrides: Partial<Record<string, unknown>> = {},
) {
	return makeExecuteMock({
		items: ENVELOPE_INPUT,
		nodeParams: (name: string, _i: number, fallback?: unknown) => {
			const map: Record<string, unknown> = {
				authKind: 'classicBot',
				resource: 'cardState',
				operation: 'select',
				conversationSource: 'envelope',
				stateKey,
				variants: { values: variants },
				options: { fallbackText: 'Card' },
				...overrides,
			};
			return name in map ? map[name] : fallback;
		},
		credentialName: 'm365AgentApi',
		credentials: { appType: 'SingleTenant', clientId: 'c', tenantId: 't', clientSecret: 's' },
	});
}

describe('cardState/select operation', () => {
	const variants: Variant[] = [
		{
			key: 'open',
			cardJson:
				'{"type":"AdaptiveCard","version":"1.5","body":[{"type":"TextBlock","text":"Open: ${title}"}]}',
			bindingData: '{"title":"Order #1"}',
		},
		{
			key: 'closed',
			cardJson:
				'{"type":"AdaptiveCard","version":"1.5","body":[{"type":"TextBlock","text":"Closed"}]}',
		},
	];

	beforeEach(() => {
		sendToConversation.mockClear();
		sendToConversation.mockResolvedValue({ id: 'act-cardstate-send' });
	});

	it('picks variant by state key and sends bound card', async () => {
		const ctx = makeCtxFor('open', variants);
		const bundles = new Map<string, BotConnectorBundle>();
		await cardStateSelectExecute.call(
			ctx,
			0,
			'classicBot' as any,
			{ appType: 'SingleTenant', clientId: 'c' } as any,
			bundles,
		);
		expect(sendToConversation).toHaveBeenCalledTimes(1);
		const activity = sendToConversation.mock.calls[0][1];
		const cardContent = activity.attachments?.[0]?.content;
		expect(cardContent.body[0].text).toBe('Open: Order #1');
	});

	it('variant without bindingData renders raw JSON', async () => {
		const ctx = makeCtxFor('closed', variants);
		const bundles = new Map<string, BotConnectorBundle>();
		await cardStateSelectExecute.call(
			ctx,
			0,
			'classicBot' as any,
			{ appType: 'SingleTenant', clientId: 'c' } as any,
			bundles,
		);
		const activity = sendToConversation.mock.calls[0][1];
		const cardContent = activity.attachments?.[0]?.content;
		expect(cardContent.body[0].text).toBe('Closed');
	});

	it('unknown state key throws NodeOperationError naming available keys', async () => {
		const ctx = makeCtxFor('unknown', variants);
		const bundles = new Map<string, BotConnectorBundle>();
		await expect(
			cardStateSelectExecute.call(
				ctx,
				0,
				'classicBot' as any,
				{ appType: 'SingleTenant', clientId: 'c' } as any,
				bundles,
			),
		).rejects.toThrow(/stateKey.*open.*closed/);
	});

	it('invalid card JSON in a variant throws NodeOperationError with field name', async () => {
		const bad: Variant[] = [{ key: 'x', cardJson: '{not json' }];
		const ctx = makeCtxFor('x', bad);
		const bundles = new Map<string, BotConnectorBundle>();
		await expect(
			cardStateSelectExecute.call(
				ctx,
				0,
				'classicBot' as any,
				{ appType: 'SingleTenant', clientId: 'c' } as any,
				bundles,
			),
		).rejects.toThrow(/variants.*cardJson/i);
	});
});
