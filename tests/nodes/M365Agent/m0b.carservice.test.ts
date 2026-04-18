import { describe, it, expect, vi, beforeEach } from 'vitest';
import { M365Agent } from '../../../nodes/M365Agent/M365Agent.node';
import { makeExecuteContext, makeCredentials } from '../../helpers/makeContext';

// ---------------------------------------------------------------------------
// Mock the auth router and bot connector — no real network calls
// ---------------------------------------------------------------------------

vi.mock('../../../shared/auth/router', () => ({
	acquireOutboundToken: vi.fn().mockResolvedValue({ authorizationHeader: 'Bearer fake-token' }),
}));

const fakeClient = {
	replyToActivity: vi.fn().mockResolvedValue({ id: 'new-reply-id' }),
	sendToConversation: vi.fn(),
	updateActivity: vi.fn(),
	deleteActivity: vi.fn(),
};

const fakeBundle = {
	client: fakeClient,
	axios: {},
	token: 'fake-token',
	baseURL: 'https://smba.trafficmanager.net/emea/',
};

const mockCreateConnectorFromBearer = vi.fn().mockReturnValue(fakeBundle);

vi.mock('../../../shared/botConnector', () => ({
	createConnector: vi.fn(),
	createConnectorFromBearer: (...args: unknown[]) => mockCreateConnectorFromBearer(...args),
	replyInThread: vi.fn(),
}));

// ---------------------------------------------------------------------------
// Car-service envelope — representative inbound item from M365AgentTrigger
// ---------------------------------------------------------------------------

const CAR_SERVICE_ENVELOPE = {
	conversationReference: {
		serviceUrl: 'https://smba.trafficmanager.net/emea/',
		conversation: { id: '19:carservice@thread.tacv2' },
		activityId: 'act-carservice-inbound-001',
		channelId: 'msteams',
		bot: { id: '28:bot-id', name: 'CarServiceBot' },
		user: { id: 'alice-user-id', name: 'Alice' },
		locale: 'en-US',
	},
	activity: {
		type: 'message',
		text: 'What services do you offer?',
		from: { id: 'alice-user-id', name: 'Alice' },
	},
	parsed: {
		type: 'message',
		text: 'What services do you offer?',
		userName: 'Alice',
		userId: 'alice-user-id',
	},
	// Ancillary fields — must survive round-trip per manifest §12
	teamsMessageId: 'stored-parent-id',
	customWorkflowField: 'preserved-value',
};

// ---------------------------------------------------------------------------
// M0B car-service parity test suite
// ---------------------------------------------------------------------------

describe('M0B car-service parity — Message/Reply via new M365Agent', () => {
	beforeEach(() => {
		vi.clearAllMocks();
		mockCreateConnectorFromBearer.mockReturnValue(fakeBundle);
		fakeClient.replyToActivity.mockResolvedValue({ id: 'new-reply-id' });
	});

	it('routes a car-service reply through resolveConversationReference → replyToActivity', async () => {
		const node = new M365Agent();
		const ctx = makeExecuteContext({
			inputItems: [CAR_SERVICE_ENVELOPE],
			credentials: makeCredentials(),
			parameters: {
				resource: 'message',
				operation: 'reply',
				conversationSource: 'envelope',
				text: 'Hello, Alice!',
				options: {},
			},
		});

		// Cast is safe in tests — the mock satisfies the call-site subset.
		const result = await node.execute.call(
			ctx as unknown as import('n8n-workflow').IExecuteFunctions,
		);

		// createConnectorFromBearer called once with the envelope's serviceUrl and auth header
		expect(mockCreateConnectorFromBearer).toHaveBeenCalledTimes(1);
		expect(mockCreateConnectorFromBearer).toHaveBeenCalledWith(
			CAR_SERVICE_ENVELOPE.conversationReference.serviceUrl,
			'Bearer fake-token',
		);

		// replyToActivity called once with (conversationId, activityId, activity)
		expect(fakeClient.replyToActivity).toHaveBeenCalledTimes(1);
		const [convId, actId, activity] = fakeClient.replyToActivity.mock.calls[0];
		expect(convId).toBe(CAR_SERVICE_ENVELOPE.conversationReference.conversation.id);
		expect(actId).toBe(CAR_SERVICE_ENVELOPE.conversationReference.activityId);
		expect(activity.type).toBe('message');
		expect(activity.text).toBe('Hello, Alice!');

		// Output shape: one branch, one item, envelope preserved, replyResult added, pairedItem=0
		expect(result).toHaveLength(1); // first output branch
		expect(result[0]).toHaveLength(1); // one item
		const outItem = result[0][0];
		expect(outItem.pairedItem).toBe(0);
		expect(outItem.json).toMatchObject({
			// Input envelope preserved (manifest §12)
			conversationReference: CAR_SERVICE_ENVELOPE.conversationReference,
			activity: CAR_SERVICE_ENVELOPE.activity,
			parsed: CAR_SERVICE_ENVELOPE.parsed,
			teamsMessageId: 'stored-parent-id',
			customWorkflowField: 'preserved-value',
			// New field added by M365Agent reply operation
			replyResult: { id: 'new-reply-id' },
		});
	});
});

describe('Message/Send with mentions + suggested actions', () => {
	beforeEach(() => {
		vi.clearAllMocks();
		mockCreateConnectorFromBearer.mockReturnValue(fakeBundle);
		fakeClient.sendToConversation.mockResolvedValue({ id: 'new-send-id' });
	});

	it('prepends <at>Alice</at> token + attaches mention entity + suggestedActions', async () => {
		const node = new M365Agent();
		const ctx = makeExecuteContext({
			inputItems: [
				{
					conversationReference: {
						serviceUrl: 'https://smba.trafficmanager.net/emea/',
						conversation: { id: '19:x@thread.tacv2' },
						channelId: 'msteams',
					},
					activity: {},
				},
			],
			credentials: makeCredentials(),
			parameters: {
				resource: 'message',
				operation: 'send',
				conversationSource: 'envelope',
				text: 'hello there',
				options: {
					mentions: {
						values: [{ type: 'user', id: '29:abc', name: 'Alice' }],
					},
					suggestedActions: {
						values: [
							{ type: 'imBack', title: 'Yes', value: 'yes' },
							{ type: 'imBack', title: 'No', value: 'no' },
						],
					},
				},
			},
		});

		await node.execute.call(ctx as unknown as import('n8n-workflow').IExecuteFunctions);

		expect(fakeClient.sendToConversation).toHaveBeenCalledTimes(1);
		const [, activity] = fakeClient.sendToConversation.mock.calls[0];
		expect(activity.text).toBe('<at>Alice</at> hello there');
		expect(activity.entities).toEqual([
			{ type: 'mention', mentioned: { id: '29:abc', name: 'Alice' }, text: '<at>Alice</at>' },
		]);
		expect(activity.suggestedActions?.actions).toHaveLength(2);
		expect(activity.suggestedActions?.actions[0]).toEqual({
			type: 'imBack',
			title: 'Yes',
			value: 'yes',
		});
	});
});
