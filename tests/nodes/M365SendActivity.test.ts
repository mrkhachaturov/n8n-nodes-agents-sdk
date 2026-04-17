import { describe, it, expect, vi, beforeEach } from 'vitest';
import { M365SendActivity } from '../../nodes/M365SendActivity/M365SendActivity.node';
import { makeExecuteContext, makeCredentials } from '../helpers/makeContext';

// ---------------------------------------------------------------------------
// Mock the bot connector — no real network calls in unit tests
// ---------------------------------------------------------------------------

const fakeClient = {
	replyToActivity: vi.fn().mockResolvedValue({ id: 'new-reply-id' }),
	sendToConversation: vi.fn().mockResolvedValue({ id: 'new-proactive-id' }),
	updateActivity: vi.fn().mockResolvedValue({ id: 'updated-id' }),
	deleteActivity: vi.fn().mockResolvedValue(undefined),
};

const fakeBundle = {
	client: fakeClient,
	axios: {},
	token: 'fake-token',
	baseURL: 'https://smba.trafficmanager.net/amer/',
};

const mockCreateConnector = vi.fn().mockResolvedValue(fakeBundle);

vi.mock('../../shared/botConnector', () => ({
	createConnector: (...args: unknown[]) => mockCreateConnector(...args),
	replyInThread: vi.fn().mockResolvedValue({ id: 'thread-reply-id' }),
}));

describe('M365SendActivity description', () => {
	it('has all five operations', () => {
		const n = new M365SendActivity();
		const op = n.description.properties.find((p) => p.name === 'operation');
		const values = (op?.options as { value: string }[] | undefined)?.map((o) => o.value).sort();
		expect(values).toEqual(['delete', 'proactive', 'reply', 'replyInThread', 'update']);
	});

	it('exposes parentActivityId for replyInThread only', () => {
		const n = new M365SendActivity();
		const p = n.description.properties.find((x) => x.name === 'parentActivityId');
		const shown = (p?.displayOptions?.show as Record<string, unknown> | undefined)?.operation;
		expect(shown).toEqual(['replyInThread']);
	});

	it('binds to m365AgentApi credential', () => {
		const n = new M365SendActivity();
		const credName = (n.description.credentials ?? [])[0]?.name;
		expect(credName).toBe('m365AgentApi');
	});
});

// ---------------------------------------------------------------------------
// execute() behaviour tests
// ---------------------------------------------------------------------------

const BASE_REF = {
	serviceUrl: 'https://smba.trafficmanager.net/amer/',
	conversation: { id: 'conv-001' },
	activityId: 'act-001',
	channelId: 'msteams',
};

const ACTIVITY = { type: 'message', text: 'Hello!' };

describe('M365SendActivity execute()', () => {
	beforeEach(() => {
		vi.clearAllMocks();
		mockCreateConnector.mockResolvedValue(fakeBundle);
		fakeClient.replyToActivity.mockResolvedValue({ id: 'new-reply-id' });
		fakeClient.sendToConversation.mockResolvedValue({ id: 'new-proactive-id' });
		fakeClient.updateActivity.mockResolvedValue({ id: 'updated-id' });
		fakeClient.deleteActivity.mockResolvedValue(undefined);
	});

	it('reply: calls replyToActivity with conversationId, activityId, and activity', async () => {
		const ctx = makeExecuteContext({
			inputItems: [{ conversationReference: BASE_REF, activity: ACTIVITY }],
			credentials: makeCredentials(),
			parameters: {
				operation: 'reply',
				conversationReference: BASE_REF,
				activity: ACTIVITY,
			},
		});

		const node = new M365SendActivity();
		const [output] = await node.execute.call(ctx as never);

		expect(fakeClient.replyToActivity).toHaveBeenCalledWith(
			BASE_REF.conversation.id,
			BASE_REF.activityId,
			ACTIVITY,
		);
		expect(output[0].json).toMatchObject({ sendResult: { id: 'new-reply-id' } });
	});

	it('proactive: calls sendToConversation and does not require activityId', async () => {
		const refWithoutActivityId = {
			serviceUrl: BASE_REF.serviceUrl,
			conversation: BASE_REF.conversation,
			channelId: BASE_REF.channelId,
			// no activityId
		};

		const ctx = makeExecuteContext({
			inputItems: [{ conversationReference: refWithoutActivityId, activity: ACTIVITY }],
			credentials: makeCredentials(),
			parameters: {
				operation: 'proactive',
				conversationReference: refWithoutActivityId,
				activity: ACTIVITY,
			},
		});

		const node = new M365SendActivity();
		const [output] = await node.execute.call(ctx as never);

		expect(fakeClient.sendToConversation).toHaveBeenCalledWith(
			refWithoutActivityId.conversation.id,
			ACTIVITY,
		);
		expect(output[0].json).toMatchObject({ sendResult: { id: 'new-proactive-id' } });
	});

	it('reply: throws NodeOperationError when activityId is missing', async () => {
		const refWithoutActivityId = {
			serviceUrl: BASE_REF.serviceUrl,
			conversation: BASE_REF.conversation,
			channelId: BASE_REF.channelId,
		};

		const ctx = makeExecuteContext({
			inputItems: [{ conversationReference: refWithoutActivityId, activity: ACTIVITY }],
			credentials: makeCredentials(),
			parameters: {
				operation: 'reply',
				conversationReference: refWithoutActivityId,
				activity: ACTIVITY,
			},
		});

		const node = new M365SendActivity();
		await expect(node.execute.call(ctx as never)).rejects.toThrow(/activityId/);
	});

	it('delete: calls deleteActivity and does not require an activity body', async () => {
		const ctx = makeExecuteContext({
			inputItems: [{ conversationReference: BASE_REF }],
			credentials: makeCredentials(),
			parameters: {
				operation: 'delete',
				conversationReference: BASE_REF,
				// no 'activity' parameter for delete
			},
		});

		const node = new M365SendActivity();
		const [output] = await node.execute.call(ctx as never);

		expect(fakeClient.deleteActivity).toHaveBeenCalledWith(
			BASE_REF.conversation.id,
			BASE_REF.activityId,
		);
		expect(output[0].json).toMatchObject({
			sendResult: { ok: true, deletedActivityId: BASE_REF.activityId },
		});
	});

	it('memoizes createConnector: called once per unique serviceUrl across multiple items', async () => {
		const ref1 = { ...BASE_REF, conversation: { id: 'conv-A' } };
		const ref2 = { ...BASE_REF, conversation: { id: 'conv-B' } };
		// Both items share the same serviceUrl — connector should be created once
		const ctx = makeExecuteContext({
			inputItems: [
				{ conversationReference: ref1, activity: ACTIVITY },
				{ conversationReference: ref2, activity: ACTIVITY },
			],
			credentials: makeCredentials(),
			parameters: {
				operation: 'proactive',
				'conversationReference:0': ref1,
				'conversationReference:1': ref2,
				'activity:0': ACTIVITY,
				'activity:1': ACTIVITY,
			},
		});

		const node = new M365SendActivity();
		await node.execute.call(ctx as never);

		// Key assertion: one unique serviceUrl → createConnector called exactly once
		expect(mockCreateConnector).toHaveBeenCalledTimes(1);
		expect(mockCreateConnector).toHaveBeenCalledWith(expect.anything(), BASE_REF.serviceUrl);
	});
});
