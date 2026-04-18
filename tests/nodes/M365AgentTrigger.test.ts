import { describe, it, expect, vi, beforeEach } from 'vitest';
import { M365AgentTrigger } from '../../nodes/M365AgentTrigger/M365AgentTrigger.node';
import { makeWebhookContext, makeCredentials } from '../helpers/makeContext';

// Mock verifyJwt so tests don't hit real JWKS endpoints
vi.mock('../../shared/verifyJwt', () => ({
	verifyJwt: vi.fn().mockResolvedValue({ aud: 'test-client-id' }),
}));

describe('M365AgentTrigger description', () => {
	it('registers two webhooks: POST messages + GET health', () => {
		const node = new M365AgentTrigger();
		const methods = (node.description.webhooks ?? []).map((w) => w.httpMethod).sort();
		expect(methods).toEqual(['GET', 'POST']);
	});

	it('binds to m365AgentApi credential', () => {
		const node = new M365AgentTrigger();
		const credName = (node.description.credentials ?? [])[0]?.name;
		expect(credName).toBe('m365AgentApi');
	});

	it('exposes activityTypes and channelFilter properties', () => {
		const node = new M365AgentTrigger();
		const names = node.description.properties.map((p) => p.name);
		expect(names).toContain('activityTypes');
		expect(names).toContain('channelFilter');
	});

	it('does NOT expose a raw output format (envelope contract is fixed)', () => {
		const node = new M365AgentTrigger();
		const names = node.description.properties.map((p) => p.name);
		expect(names).not.toContain('outputFormat');
	});
});

// ---------------------------------------------------------------------------
// webhook() behaviour tests
// ---------------------------------------------------------------------------

describe('M365AgentTrigger webhook()', () => {
	let node: M365AgentTrigger;

	beforeEach(() => {
		node = new M365AgentTrigger();
		vi.clearAllMocks();
	});

	it('returns 200 + status ok on GET (health check)', async () => {
		const ctx = makeWebhookContext({ method: 'GET' });
		const result = await node.webhook.call(ctx as never);

		expect(ctx._statusFn).toHaveBeenCalledWith(200);
		expect(ctx._jsonFn).toHaveBeenCalledWith({ status: 'ok', service: 'M365 Agent' });
		expect(result).toEqual({ noWebhookResponse: true });
	});

	it('returns 401 when no bearer token and anonymousAllowed is false', async () => {
		const ctx = makeWebhookContext({
			method: 'POST',
			headers: {}, // no authorization header
			credentials: makeCredentials({ anonymousAllowed: false }),
		});

		const result = await node.webhook.call(ctx as never);

		expect(ctx._statusFn).toHaveBeenCalledWith(401);
		expect(ctx._jsonFn).toHaveBeenCalledWith(
			expect.objectContaining({ error: 'Missing bearer token' }),
		);
		expect(result).toEqual({ noWebhookResponse: true });
	});

	it('emits full envelope on valid POST with a message activity', async () => {
		const { verifyJwt } = await import('../../shared/verifyJwt');
		(verifyJwt as ReturnType<typeof vi.fn>).mockResolvedValueOnce({ aud: 'test-client-id' });

		const body = {
			id: 'act-001',
			type: 'message',
			serviceUrl: 'https://smba.trafficmanager.net/amer/',
			channelId: 'msteams',
			conversation: { id: 'conv-001', conversationType: 'personal' },
			from: { id: 'user-001', name: 'Alice' },
			recipient: { id: 'bot-001', name: 'TestBot' },
		};

		const ctx = makeWebhookContext({
			method: 'POST',
			headers: { authorization: 'Bearer valid.jwt.token' },
			body,
			credentials: makeCredentials({ anonymousAllowed: false }),
			parameters: { activityTypes: ['message'], channelFilter: [] },
		});

		const result = await node.webhook.call(ctx as never);

		expect(result.workflowData).toBeDefined();
		const item = result.workflowData![0][0];
		expect(item.json).toMatchObject({
			conversationReference: expect.objectContaining({
				serviceUrl: body.serviceUrl,
				activityId: body.id,
			}),
			activity: expect.objectContaining({ type: 'message' }),
			parsed: expect.objectContaining({ type: 'message' }),
			raw: expect.objectContaining({ id: body.id }),
		});
	});

	it('filters activity by type and returns empty workflowData', async () => {
		const { verifyJwt } = await import('../../shared/verifyJwt');
		(verifyJwt as ReturnType<typeof vi.fn>).mockResolvedValueOnce({ aud: 'test-client-id' });

		const body = {
			id: 'act-002',
			type: 'typing',
			serviceUrl: 'https://smba.trafficmanager.net/amer/',
			channelId: 'msteams',
			conversation: { id: 'conv-002' },
		};

		const ctx = makeWebhookContext({
			method: 'POST',
			headers: { authorization: 'Bearer valid.jwt.token' },
			body,
			credentials: makeCredentials({ anonymousAllowed: false }),
			parameters: { activityTypes: ['message'], channelFilter: [] },
		});

		const result = await node.webhook.call(ctx as never);

		expect(result.workflowData).toEqual([[]]); // filtered → empty emission
		expect(result.webhookResponse).toMatchObject({ status: 200 });
	});
});

// ---------------------------------------------------------------------------
// responseMode parameter tests
// ---------------------------------------------------------------------------

describe('M365AgentTrigger — responseMode parameter', () => {
	it('displayName is "M365 Agent Trigger"', () => {
		const node = new M365AgentTrigger();
		expect(node.description.displayName).toBe('M365 Agent Trigger');
	});

	it('exposes a responseMode property with onReceived default', () => {
		const node = new M365AgentTrigger();
		const prop = node.description.properties.find((p) => p.name === 'responseMode');
		expect(prop).toBeDefined();
		expect(prop?.default).toBe('onReceived');
		expect(prop?.noDataExpression).toBe(true);
		const values = (prop?.options as { value: string }[] | undefined)?.map((o) => o.value).sort();
		expect(values).toEqual(['onReceived', 'responseNode']);
	});

	it('POST webhook responseMode is an expression referencing the parameter', () => {
		const node = new M365AgentTrigger();
		const post = (node.description.webhooks ?? []).find((w) => w.httpMethod === 'POST');
		expect(post?.responseMode).toBe('={{$parameter["responseMode"]}}');
	});

	it('GET webhook responseMode stays onReceived', () => {
		const node = new M365AgentTrigger();
		const get = (node.description.webhooks ?? []).find((w) => w.httpMethod === 'GET');
		expect(get?.responseMode).toBe('onReceived');
	});

	it('does not expose usableAsTool (manifest §10.3 — triggers are not tools)', () => {
		expect(new M365AgentTrigger().description.usableAsTool).toBeFalsy();
	});
});
