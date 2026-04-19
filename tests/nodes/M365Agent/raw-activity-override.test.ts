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

// All 20 card executors (10 resources × send + update)
import { execute as cardAdaptiveSend } from '../../../nodes/M365Agent/actions/card-adaptive/send.operation';
import { execute as cardAdaptiveUpdate } from '../../../nodes/M365Agent/actions/card-adaptive/update.operation';
import { execute as cardAnimationSend } from '../../../nodes/M365Agent/actions/card-animation/send.operation';
import { execute as cardAnimationUpdate } from '../../../nodes/M365Agent/actions/card-animation/update.operation';
import { execute as cardAudioSend } from '../../../nodes/M365Agent/actions/card-audio/send.operation';
import { execute as cardAudioUpdate } from '../../../nodes/M365Agent/actions/card-audio/update.operation';
import { execute as cardHeroSend } from '../../../nodes/M365Agent/actions/card-hero/send.operation';
import { execute as cardHeroUpdate } from '../../../nodes/M365Agent/actions/card-hero/update.operation';
import { execute as cardO365Send } from '../../../nodes/M365Agent/actions/card-o365-connector/send.operation';
import { execute as cardO365Update } from '../../../nodes/M365Agent/actions/card-o365-connector/update.operation';
import { execute as cardRawSend } from '../../../nodes/M365Agent/actions/card-raw-attachment/send.operation';
import { execute as cardRawUpdate } from '../../../nodes/M365Agent/actions/card-raw-attachment/update.operation';
import { execute as cardReceiptSend } from '../../../nodes/M365Agent/actions/card-receipt/send.operation';
import { execute as cardReceiptUpdate } from '../../../nodes/M365Agent/actions/card-receipt/update.operation';
import { execute as cardSignInSend } from '../../../nodes/M365Agent/actions/card-sign-in/send.operation';
import { execute as cardSignInUpdate } from '../../../nodes/M365Agent/actions/card-sign-in/update.operation';
import { execute as cardThumbnailSend } from '../../../nodes/M365Agent/actions/card-thumbnail/send.operation';
import { execute as cardThumbnailUpdate } from '../../../nodes/M365Agent/actions/card-thumbnail/update.operation';
import { execute as cardVideoSend } from '../../../nodes/M365Agent/actions/card-video/send.operation';
import { execute as cardVideoUpdate } from '../../../nodes/M365Agent/actions/card-video/update.operation';

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

// ---------------------------------------------------------------------------
// Message ops matrix (preserved from Task 9)
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// Card ops matrix — 10 resources × send/update × 3 scenarios = 60 cases
// ---------------------------------------------------------------------------

interface CardOpSpec {
	name: string;
	execute: typeof cardAdaptiveSend;
	kind: 'send' | 'update';
	resource: string;
	// Required node params to satisfy each resource's buildAttachment() — merged
	// with a shared base (authKind, conversationSource, options).
	extraParams?: Record<string, unknown>;
}

// A valid Adaptive Card template (no bindings needed).
const ADAPTIVE_TEMPLATE = JSON.stringify({
	type: 'AdaptiveCard',
	version: '1.4',
	body: [{ type: 'TextBlock', text: 'hello' }],
});

// Media list to satisfy Animation / Audio / Video buildAttachment validation.
const MEDIA_PARAM = { mediaItem: [{ url: 'https://example.com/clip.mp4', options: {} }] };

const CARD_OPS: CardOpSpec[] = [
	{
		name: 'card-adaptive.send',
		execute: cardAdaptiveSend,
		kind: 'send',
		resource: 'adaptiveCard',
		extraParams: { cardTemplate: ADAPTIVE_TEMPLATE, bindingData: {} },
	},
	{
		name: 'card-adaptive.update',
		execute: cardAdaptiveUpdate,
		kind: 'update',
		resource: 'adaptiveCard',
		extraParams: { cardTemplate: ADAPTIVE_TEMPLATE, bindingData: {} },
	},
	{
		name: 'card-animation.send',
		execute: cardAnimationSend,
		kind: 'send',
		resource: 'animationCard',
		extraParams: { title: 'Anim', media: MEDIA_PARAM },
	},
	{
		name: 'card-animation.update',
		execute: cardAnimationUpdate,
		kind: 'update',
		resource: 'animationCard',
		extraParams: { title: 'Anim', media: MEDIA_PARAM },
	},
	{
		name: 'card-audio.send',
		execute: cardAudioSend,
		kind: 'send',
		resource: 'audioCard',
		extraParams: { title: 'Audio', media: MEDIA_PARAM },
	},
	{
		name: 'card-audio.update',
		execute: cardAudioUpdate,
		kind: 'update',
		resource: 'audioCard',
		extraParams: { title: 'Audio', media: MEDIA_PARAM },
	},
	{
		name: 'card-hero.send',
		execute: cardHeroSend,
		kind: 'send',
		resource: 'heroCard',
		extraParams: { title: 'Hero', text: 'body' },
	},
	{
		name: 'card-hero.update',
		execute: cardHeroUpdate,
		kind: 'update',
		resource: 'heroCard',
		extraParams: { title: 'Hero', text: 'body' },
	},
	{
		name: 'card-o365-connector.send',
		execute: cardO365Send,
		kind: 'send',
		resource: 'o365ConnectorCard',
		extraParams: { cardContent: JSON.stringify({ title: 'T', text: 'x' }) },
	},
	{
		name: 'card-o365-connector.update',
		execute: cardO365Update,
		kind: 'update',
		resource: 'o365ConnectorCard',
		extraParams: { cardContent: JSON.stringify({ title: 'T', text: 'x' }) },
	},
	{
		name: 'card-raw-attachment.send',
		execute: cardRawSend,
		kind: 'send',
		resource: 'rawAttachment',
		extraParams: {
			contentType: 'application/vnd.microsoft.card.custom',
			content: JSON.stringify({ foo: 'bar' }),
		},
	},
	{
		name: 'card-raw-attachment.update',
		execute: cardRawUpdate,
		kind: 'update',
		resource: 'rawAttachment',
		extraParams: {
			contentType: 'application/vnd.microsoft.card.custom',
			content: JSON.stringify({ foo: 'bar' }),
		},
	},
	{
		name: 'card-receipt.send',
		execute: cardReceiptSend,
		kind: 'send',
		resource: 'receiptCard',
		extraParams: { cardContent: JSON.stringify({ title: 'Receipt', items: [] }) },
	},
	{
		name: 'card-receipt.update',
		execute: cardReceiptUpdate,
		kind: 'update',
		resource: 'receiptCard',
		extraParams: { cardContent: JSON.stringify({ title: 'Receipt', items: [] }) },
	},
	{
		name: 'card-sign-in.send',
		execute: cardSignInSend,
		kind: 'send',
		resource: 'signInCard',
		extraParams: { title: 'Sign in', url: 'https://auth.example.com' },
	},
	{
		name: 'card-sign-in.update',
		execute: cardSignInUpdate,
		kind: 'update',
		resource: 'signInCard',
		extraParams: { title: 'Sign in', url: 'https://auth.example.com' },
	},
	{
		name: 'card-thumbnail.send',
		execute: cardThumbnailSend,
		kind: 'send',
		resource: 'thumbnailCard',
		extraParams: { title: 'Thumb', text: 'body' },
	},
	{
		name: 'card-thumbnail.update',
		execute: cardThumbnailUpdate,
		kind: 'update',
		resource: 'thumbnailCard',
		extraParams: { title: 'Thumb', text: 'body' },
	},
	{
		name: 'card-video.send',
		execute: cardVideoSend,
		kind: 'send',
		resource: 'videoCard',
		extraParams: { title: 'Video', media: MEDIA_PARAM },
	},
	{
		name: 'card-video.update',
		execute: cardVideoUpdate,
		kind: 'update',
		resource: 'videoCard',
		extraParams: { title: 'Video', media: MEDIA_PARAM },
	},
];

function buildCardNodeParams(spec: CardOpSpec, rawOverride: string | undefined) {
	const options = { rawActivityOverride: rawOverride ?? '' };
	return (name: string, _i: number, fallback?: unknown) => {
		const base: Record<string, unknown> = {
			authKind: 'classicBot',
			resource: spec.resource,
			operation: spec.kind,
			conversationSource: 'envelope',
			options,
			...(spec.extraParams ?? {}),
		};
		return name in base ? base[name] : fallback;
	};
}

function makeCardCtx(spec: CardOpSpec, rawOverride: string | undefined) {
	return makeExecuteMock({
		items: ENVELOPE_INPUT,
		nodeParams: buildCardNodeParams(spec, rawOverride),
		credentialName: 'm365AgentApi',
		credentials: { appType: 'SingleTenant', clientId: 'c', tenantId: 't', clientSecret: 's' },
	});
}

async function runCardOp(spec: CardOpSpec, ctx: ReturnType<typeof makeCardCtx>) {
	const bundles = new Map<string, BotConnectorBundle>();
	return spec.execute.call(
		ctx,
		0,
		'classicBot' as any,
		{ appType: 'SingleTenant', clientId: 'c' } as any,
		bundles,
	);
}

describe('Raw Activity Override — Card operations', () => {
	beforeEach(() => {
		mocks.sendToConversation.mockClear();
		mocks.updateActivity.mockClear();
	});

	for (const spec of CARD_OPS) {
		// Card send ops use sendToConversation(convId, activity) — activity at index 1.
		// Card update ops use updateActivity(convId, activityId, activity) — activity at index 2.
		const dispatcher = spec.kind === 'send' ? mocks.sendToConversation : mocks.updateActivity;
		const activityArgIndex = spec.kind === 'send' ? 1 : 2;

		it(`${spec.name}: empty override → constructed attachment preserved`, async () => {
			const ctx = makeCardCtx(spec, '');
			await runCardOp(spec, ctx);
			const call = dispatcher.mock.calls[0];
			expect(call).toBeDefined();
			const activity = call[activityArgIndex] as Record<string, unknown>;
			expect(activity.type).toBe('message');
			expect(Array.isArray(activity.attachments)).toBe(true);
			expect((activity.attachments as unknown[]).length).toBeGreaterThan(0);
		});

		it(`${spec.name}: populated override → override body dispatched as-is`, async () => {
			const override = JSON.stringify({
				type: 'message',
				text: 'full override',
				channelData: { overridden: true },
			});
			const ctx = makeCardCtx(spec, override);
			await runCardOp(spec, ctx);
			const call = dispatcher.mock.calls[0];
			expect(call).toBeDefined();
			const activity = call[activityArgIndex] as Record<string, unknown>;
			expect(activity.text).toBe('full override');
			expect(activity.channelData).toEqual({ overridden: true });
			// Constructed attachments are discarded when the override replaces the body.
			expect(activity.attachments).toBeUndefined();
		});

		it(`${spec.name}: invalid JSON in override → NodeOperationError with field name`, async () => {
			const ctx = makeCardCtx(spec, '{not json');
			await expect(runCardOp(spec, ctx)).rejects.toThrow(/rawActivityOverride/i);
		});
	}
});
