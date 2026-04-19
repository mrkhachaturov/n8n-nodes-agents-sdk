import { describe, it, expect, vi, beforeEach } from 'vitest';

// Shared mocks — use vi.hoisted so the same vi.fn instances are visible both
// inside the hoisted vi.mock factories and in the test assertions below.
const mocks = vi.hoisted(() => ({
	sendToConversation: vi.fn().mockResolvedValue({ id: 'act-send' }),
	replyToActivity: vi.fn().mockResolvedValue({ id: 'act-reply' }),
	updateActivity: vi.fn().mockResolvedValue({ id: 'act-update' }),
	replyInThread: vi.fn().mockResolvedValue({ id: 'act-thread' }),
}));

vi.mock('../../../shared/auth/router', () => ({
	acquireOutboundToken: vi.fn().mockResolvedValue({ authorizationHeader: 'Bearer routed' }),
}));
vi.mock('../../../shared/botConnector', () => ({
	createConnectorFromBearer: vi.fn().mockImplementation((url: string, header: string) => ({
		client: {
			sendToConversation: mocks.sendToConversation,
			replyToActivity: mocks.replyToActivity,
			updateActivity: mocks.updateActivity,
		},
		axios: { post: vi.fn().mockResolvedValue({ data: { id: 'act-thread' } }) },
		token: header,
		baseURL: url,
	})),
	replyInThread: mocks.replyInThread,
}));

import { execute as sendExecute } from '../../../nodes/M365Agent/actions/message/send.operation';
import { execute as replyExecute } from '../../../nodes/M365Agent/actions/message/reply.operation';
import { execute as updateExecute } from '../../../nodes/M365Agent/actions/message/update.operation';
import { execute as replyInThreadExecute } from '../../../nodes/M365Agent/actions/message/replyInThread.operation';
import type { BotConnectorBundle } from '../../../shared/botConnector';
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
		},
	},
];

type OpName = 'send' | 'reply' | 'update' | 'replyInThread';

interface OpSpec {
	name: OpName;
	operation: string;
	execute: typeof sendExecute;
	dispatcher: ReturnType<typeof vi.fn>;
	// Index of the activity argument in the mock call.
	activityArgIndex: number;
}

const OPS: OpSpec[] = [
	{
		name: 'send',
		operation: 'send',
		execute: sendExecute,
		dispatcher: mocks.sendToConversation,
		activityArgIndex: 1,
	},
	{
		name: 'reply',
		operation: 'reply',
		execute: replyExecute,
		dispatcher: mocks.replyToActivity,
		activityArgIndex: 2,
	},
	{
		name: 'update',
		operation: 'update',
		execute: updateExecute,
		dispatcher: mocks.updateActivity,
		activityArgIndex: 2,
	},
	{
		name: 'replyInThread',
		operation: 'replyInThread',
		execute: replyInThreadExecute,
		// replyInThread is called as replyInThread(bundle, convId, parentActivityId, activity) — activity at index 3.
		dispatcher: mocks.replyInThread,
		activityArgIndex: 3,
	},
];

function buildNodeParams(spec: OpSpec, rawOverride: string | undefined) {
	return (name: string, _i: number, fallback?: unknown) => {
		const map: Record<string, unknown> = {
			authKind: 'classicBot',
			resource: 'message',
			operation: spec.operation,
			conversationSource: 'envelope',
			text: 'constructed body',
			parentActivityId: 'parent-1',
			options: { rawActivityOverride: rawOverride ?? '' },
		};
		return map[name] ?? fallback;
	};
}

function makeCtx(spec: OpSpec, rawOverride: string | undefined) {
	return makeExecuteMock({
		items: ENVELOPE_INPUT,
		nodeParams: buildNodeParams(spec, rawOverride),
		credentialName: 'm365AgentApi',
		credentials: { appType: 'SingleTenant', clientId: 'c', tenantId: 't', clientSecret: 's' },
	});
}

async function runOp(spec: OpSpec, ctx: ReturnType<typeof makeCtx>) {
	const bundles = new Map<string, BotConnectorBundle>();
	return spec.execute.call(
		ctx,
		0,
		'classicBot' as any,
		{ appType: 'SingleTenant', clientId: 'c' } as any,
		bundles,
	);
}

describe('Raw Activity Override — Message operations', () => {
	beforeEach(() => {
		mocks.sendToConversation.mockClear();
		mocks.replyToActivity.mockClear();
		mocks.updateActivity.mockClear();
		mocks.replyInThread.mockClear();
	});

	for (const spec of OPS) {
		it(`${spec.name}: empty override → constructed body sent`, async () => {
			const ctx = makeCtx(spec, '');
			await runOp(spec, ctx);
			const call = spec.dispatcher.mock.calls[0];
			expect(call).toBeDefined();
			const activity = call[spec.activityArgIndex] as Record<string, unknown>;
			expect(activity.text).toBe('constructed body');
			expect(activity.customChannelData).toBeUndefined();
		});

		it(`${spec.name}: populated override → override body sent`, async () => {
			const override = JSON.stringify({
				type: 'message',
				text: 'from override',
				customChannelData: true,
			});
			const ctx = makeCtx(spec, override);
			await runOp(spec, ctx);
			const call = spec.dispatcher.mock.calls[0];
			expect(call).toBeDefined();
			const activity = call[spec.activityArgIndex] as Record<string, unknown>;
			expect(activity.text).toBe('from override');
			expect(activity.customChannelData).toBe(true);
		});

		it(`${spec.name}: invalid JSON in override → NodeOperationError with field name`, async () => {
			const ctx = makeCtx(spec, '{not json');
			await expect(runOp(spec, ctx)).rejects.toThrow(/rawActivityOverride/i);
		});
	}
});
