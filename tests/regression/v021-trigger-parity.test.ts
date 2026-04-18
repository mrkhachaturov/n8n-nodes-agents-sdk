import { describe, it, expect, vi, beforeEach } from 'vitest';
import fixtureActivity from '../fixtures/v021-classic-activity.json';
import { M365AgentTrigger } from '../../nodes/M365AgentTrigger/M365AgentTrigger.node';
import { makeWebhookMock } from '../helpers/webhookMock';

vi.mock('../../shared/auth/router', () => ({
	validateInboundToken: vi.fn(async () => ({
		claims: { iss: 'https://api.botframework.com', aud: 'test-bot-app-id' },
	})),
}));

describe('v0.2.1 classic-path trigger parity', () => {
	beforeEach(() => vi.clearAllMocks());

	it('emits envelope with no authContext when authKind=classicBot', async () => {
		const ctx = makeWebhookMock({
			body: fixtureActivity,
			headers: { authorization: 'Bearer signed.jwt.stub' },
			nodeParams: {
				authKind: 'classicBot',
				responseMode: 'immediate',
				activityTypes: [],
				channelFilter: [],
			},
			credentialName: 'm365AgentApi',
			credentials: { appType: 'SingleTenant', clientId: 'cid', clientSecret: 's', tenantId: 'tid' },
		});
		const trigger = new M365AgentTrigger();
		const result = await trigger.webhook!.call(ctx);
		const envelope = (result as any).workflowData?.[0]?.[0]?.json as Record<string, unknown>;
		expect(envelope).toBeDefined();
		expect(envelope.authContext).toBeUndefined();
		// Envelope must contain the nested activity and conversationReference
		expect(envelope.activity).toBeDefined();
		expect((envelope.activity as any).type).toBe('message');
		expect(envelope.conversationReference).toBeDefined();
	});

	it('preserves v0.2.1 raw activity shape byte-for-byte', async () => {
		const ctx = makeWebhookMock({
			body: fixtureActivity,
			headers: { authorization: 'Bearer signed.jwt.stub' },
			nodeParams: {
				authKind: 'classicBot',
				responseMode: 'immediate',
				activityTypes: [],
				channelFilter: [],
			},
			credentialName: 'm365AgentApi',
			credentials: { appType: 'SingleTenant', clientId: 'cid', clientSecret: 's', tenantId: 'tid' },
		});
		const trigger = new M365AgentTrigger();
		const result = await trigger.webhook!.call(ctx);
		const envelope = (result as any).workflowData?.[0]?.[0]?.json as Record<string, unknown>;
		expect(envelope).toBeDefined();
		const activity = envelope.activity as Record<string, unknown>;
		// Assert all fixture fields are present and unmodified
		expect(activity.type).toBe(fixtureActivity.type);
		expect(activity.id).toBe(fixtureActivity.id);
		expect(activity.timestamp).toBe(fixtureActivity.timestamp);
		expect(activity.serviceUrl).toBe(fixtureActivity.serviceUrl);
		expect(activity.channelId).toBe(fixtureActivity.channelId);
		expect(activity.text).toBe(fixtureActivity.text);
		expect(activity.locale).toBe(fixtureActivity.locale);
		// Nested objects must match
		expect(activity.from).toEqual(fixtureActivity.from);
		expect(activity.conversation).toEqual(fixtureActivity.conversation);
		expect(activity.recipient).toEqual(fixtureActivity.recipient);
		// raw field mirrors activity
		const raw = envelope.raw as Record<string, unknown>;
		expect(raw).toBeDefined();
		expect(raw.type).toBe(fixtureActivity.type);
		expect(raw.text).toBe(fixtureActivity.text);
	});
});
