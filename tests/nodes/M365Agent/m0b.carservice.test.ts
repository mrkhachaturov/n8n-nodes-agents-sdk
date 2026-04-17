import { describe, it, expect, vi, beforeEach } from 'vitest';
import { M365Agent } from '../../../nodes/M365Agent/M365Agent.node';
import { makeExecuteContext, makeCredentials } from '../../helpers/makeContext';

// ---------------------------------------------------------------------------
// Mock the bot connector — no real network calls
// ---------------------------------------------------------------------------

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

const mockCreateConnector = vi.fn().mockResolvedValue(fakeBundle);

vi.mock('../../../shared/botConnector', () => ({
	createConnector: (...args: unknown[]) => mockCreateConnector(...args),
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
		mockCreateConnector.mockResolvedValue(fakeBundle);
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
		const result = await node.execute.call(ctx as unknown as import('n8n-workflow').IExecuteFunctions);

		// createConnector called once with the envelope's serviceUrl
		expect(mockCreateConnector).toHaveBeenCalledTimes(1);
		expect(mockCreateConnector).toHaveBeenCalledWith(
			expect.anything(), // AuthConfiguration
			CAR_SERVICE_ENVELOPE.conversationReference.serviceUrl,
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
