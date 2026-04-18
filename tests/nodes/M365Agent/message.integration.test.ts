import { describe, it, expect, vi, beforeEach } from 'vitest';
import { M365Agent } from '../../../nodes/M365Agent/M365Agent.node';
import { makeExecuteContext, makeCredentials } from '../../helpers/makeContext';
// Import the mocked `replyInThread` symbol directly so we can drive it
// without a fresh require() — consistent with the existing test's import
// pattern and with Vitest's hoisted `vi.mock()` at the top of this file.
import { replyInThread as mockReplyInThread } from '../../../shared/botConnector';

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
// Proactive envelope — representative inbound item from M365AgentTrigger
// ---------------------------------------------------------------------------

const PROACTIVE_ENVELOPE = {
	conversationReference: {
		serviceUrl: 'https://smba.trafficmanager.net/emea/',
		conversation: { id: '19:sample@thread.tacv2' },
		activityId: 'act-sample-inbound-001',
		channelId: 'msteams',
		bot: { id: '28:bot-id', name: 'SampleBot' },
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
// Message resource — Reply with proactive (envelope-driven) activity
// ---------------------------------------------------------------------------

describe('Message resource — Reply with proactive (envelope-driven) activity', () => {
	beforeEach(() => {
		vi.clearAllMocks();
		mockCreateConnectorFromBearer.mockReturnValue(fakeBundle);
		fakeClient.replyToActivity.mockResolvedValue({ id: 'new-reply-id' });
	});

	it('routes a reply through resolveConversationReference → replyToActivity', async () => {
		const node = new M365Agent();
		const ctx = makeExecuteContext({
			inputItems: [PROACTIVE_ENVELOPE],
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
			PROACTIVE_ENVELOPE.conversationReference.serviceUrl,
			'Bearer fake-token',
		);

		// replyToActivity called once with (conversationId, activityId, activity)
		expect(fakeClient.replyToActivity).toHaveBeenCalledTimes(1);
		const [convId, actId, activity] = fakeClient.replyToActivity.mock.calls[0];
		expect(convId).toBe(PROACTIVE_ENVELOPE.conversationReference.conversation.id);
		expect(actId).toBe(PROACTIVE_ENVELOPE.conversationReference.activityId);
		expect(activity.type).toBe('message');
		expect(activity.text).toBe('Hello, Alice!');

		// Output shape: one branch, one item, envelope preserved, replyResult added, pairedItem=0
		expect(result).toHaveLength(1); // first output branch
		expect(result[0]).toHaveLength(1); // one item
		const outItem = result[0][0];
		expect(outItem.pairedItem).toBe(0);
		expect(outItem.json).toMatchObject({
			// Input envelope preserved (manifest §12)
			conversationReference: PROACTIVE_ENVELOPE.conversationReference,
			activity: PROACTIVE_ENVELOPE.activity,
			parsed: PROACTIVE_ENVELOPE.parsed,
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

describe('Message/Reply with mentions + suggested actions', () => {
	beforeEach(() => {
		vi.clearAllMocks();
		mockCreateConnectorFromBearer.mockReturnValue(fakeBundle);
		fakeClient.replyToActivity.mockResolvedValue({ id: 'new-reply-id' });
	});

	it('attaches mention entity + token AND suggestedActions to the reply', async () => {
		const node = new M365Agent();
		const ctx = makeExecuteContext({
			inputItems: [PROACTIVE_ENVELOPE],
			credentials: makeCredentials(),
			parameters: {
				resource: 'message',
				operation: 'reply',
				conversationSource: 'envelope',
				text: 'thanks!',
				options: {
					mentions: { values: [{ type: 'user', id: 'alice-user-id', name: 'Alice' }] },
					suggestedActions: {
						values: [{ type: 'imBack', title: 'OK', value: 'ok' }],
					},
				},
			},
		});

		await node.execute.call(ctx as unknown as import('n8n-workflow').IExecuteFunctions);

		const [, , activity] = fakeClient.replyToActivity.mock.calls[0];
		expect(activity.text).toBe('<at>Alice</at> thanks!');
		expect(activity.entities?.[0]?.type).toBe('mention');
		expect(activity.suggestedActions?.actions).toHaveLength(1);
		expect(activity.suggestedActions?.actions[0]).toEqual({
			type: 'imBack',
			title: 'OK',
			value: 'ok',
		});
	});
});

describe('Message/Update with mentions + suggested actions', () => {
	beforeEach(() => {
		vi.clearAllMocks();
		mockCreateConnectorFromBearer.mockReturnValue(fakeBundle);
		fakeClient.updateActivity.mockResolvedValue({ id: 'updated-id' });
	});

	it('updateActivity body includes <at> token + mention entity + suggestedActions', async () => {
		const node = new M365Agent();
		const ctx = makeExecuteContext({
			inputItems: [PROACTIVE_ENVELOPE],
			credentials: makeCredentials(),
			parameters: {
				resource: 'message',
				operation: 'update',
				conversationSource: 'envelope',
				text: 'updated!',
				options: {
					mentions: { values: [{ type: 'user', id: 'alice-user-id', name: 'Alice' }] },
					suggestedActions: {
						values: [{ type: 'openUrl', title: 'Details', value: 'https://example.com/x' }],
					},
				},
			},
		});

		await node.execute.call(ctx as unknown as import('n8n-workflow').IExecuteFunctions);
		const [, , activity] = fakeClient.updateActivity.mock.calls[0];
		expect(activity.text).toBe('<at>Alice</at> updated!');
		expect(activity.entities?.[0]?.type).toBe('mention');
		expect(activity.suggestedActions?.actions).toHaveLength(1);
		expect(activity.suggestedActions?.actions[0].type).toBe('openUrl');
	});
});

describe('Message/replyInThread with mentions + suggested actions', () => {
	beforeEach(() => {
		vi.clearAllMocks();
		mockCreateConnectorFromBearer.mockReturnValue(fakeBundle);
		(mockReplyInThread as unknown as ReturnType<typeof vi.fn>).mockResolvedValue({
			id: 'new-thread-id',
		});
	});

	it('thread activity carries mention token + entity + suggestedActions AND uses the node parameter for parentActivityId', async () => {
		const node = new M365Agent();
		const ctx = makeExecuteContext({
			inputItems: [{ ...PROACTIVE_ENVELOPE, teamsMessageId: 'wrong-id-from-item' }],
			credentials: makeCredentials(),
			parameters: {
				resource: 'message',
				operation: 'replyInThread',
				conversationSource: 'envelope',
				text: 'ping',
				parentActivityId: 'correct-id-from-parameter',
				options: {
					mentions: { values: [{ type: 'everyone', name: 'Everyone' }] },
					suggestedActions: {
						values: [{ type: 'messageBack', title: 'Ack', value: 'ack', displayText: '👍' }],
					},
				},
			},
		});

		await node.execute.call(ctx as unknown as import('n8n-workflow').IExecuteFunctions);

		// replyInThread(bundle, conversationId, parentActivityId, activity)
		const [, , parentActivityIdArg, activity] = (
			mockReplyInThread as unknown as ReturnType<typeof vi.fn>
		).mock.calls[0];
		expect(parentActivityIdArg).toBe('correct-id-from-parameter');
		expect(activity.text).toBe('<at>Everyone</at> ping');
		expect(activity.entities?.[0]?.mentioned).toEqual({ id: '29:allchannel', name: 'Everyone' });
		expect(activity.suggestedActions?.actions).toHaveLength(1);
		expect(activity.suggestedActions?.actions[0].type).toBe('messageBack');
	});
});
