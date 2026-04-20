import { describe, it, expect } from 'vitest';
import { M365AgentTrigger } from '../../../nodes/M365AgentTrigger/M365AgentTrigger.node';

/**
 * The filter is implemented in webhook(). We stub IWebhookFunctions
 * and exercise each of the 18 ActivityTypes enum values.
 *
 * As of Task 2 (D1-D, spec §8), the flat `activityTypes` property has been
 * replaced by a Simple/Advanced mode toggle (manifest §6.3). Tests exercise
 * the Advanced path (all 18 types) because that's what the enum coverage
 * assertion protects; the Simple-mode gate is validated separately below.
 */
const ALL_18_TYPES = [
	'message', 'contactRelationUpdate', 'conversationUpdate', 'typing',
	'endOfConversation', 'event', 'invoke', 'invokeResponse', 'deleteUserData',
	'messageUpdate', 'messageDelete', 'installationUpdate', 'messageReaction',
	'suggestion', 'trace', 'handoff', 'command', 'commandResult',
];

type FilterMode = 'simple' | 'advanced';

function makeTriggerContext(
	configuredFilter: string[],
	inboundType: string,
	mode: FilterMode = 'advanced',
) {
	// Minimal-but-valid Activity skeleton — envelope builder requires serviceUrl, channelId, conversation.
	// Filter logic runs before envelope construction, so "skipped" cases never touch these fields.
	const body = {
		type: inboundType,
		id: 'a1',
		text: 'x',
		serviceUrl: 'https://smba/',
		channelId: 'msteams',
		from: { id: '29:u1' },
		conversation: { id: '19:c1', conversationType: 'personal' },
		recipient: { id: '28:b1' },
	};
	return {
		getRequestObject: () => ({
			headers: {},
			body,
			method: 'POST',
		}),
		getBodyData: () => body,
		getHeaderData: () => ({}),
		getNodeParameter: (name: string) => {
			if (name === 'activityFilterMode') return mode;
			if (name === 'activityTypesSimple') return mode === 'simple' ? configuredFilter : [];
			if (name === 'activityTypesAdvanced') return mode === 'advanced' ? configuredFilter : [];
			if (name === 'invokeNamesSimple') return [];
			if (name === 'invokeNamesAdvanced') return [];
			if (name === 'channelFilter') return [];
			if (name === 'authKind') return 'classicBot';
			if (name === 'responseMode') return 'onReceived';
			return undefined;
		},
		getCredentials: async () => ({ clientId: 'c', tenantId: 't', anonymousAllowed: true }),
		helpers: { returnJsonArray: (x: any) => x.map((json: any) => ({ json })) },
		getNode: () => ({ name: 'test' }),
	} as any;
}

describe('M365AgentTrigger activity-type filter — all 18 types (Advanced mode)', () => {
	const trigger = new M365AgentTrigger();

	for (const type of ALL_18_TYPES) {
		it(`accepts ${type} when filter includes it`, async () => {
			const ctx = makeTriggerContext([type], type, 'advanced');
			const res = await trigger.webhook.call(ctx as any);
			expect(res.workflowData?.[0]?.length ?? 0).toBeGreaterThan(0);
		});

		it(`skips ${type} when filter does not include it`, async () => {
			const otherType = type === 'message' ? 'typing' : 'message';
			const ctx = makeTriggerContext([otherType], type, 'advanced');
			const res = await trigger.webhook.call(ctx as any);
			expect(res.workflowData?.[0]?.length ?? 0).toBe(0);
		});
	}

	it('accepts any type when filter is empty (default = no filter)', async () => {
		const ctx = makeTriggerContext([], 'handoff', 'advanced');
		const res = await trigger.webhook.call(ctx as any);
		expect(res.workflowData?.[0]?.length ?? 0).toBeGreaterThan(0);
	});

	// Metadata assertion — protects against the UI drifting from the authoritative enum.
	it('Trigger `activityTypesAdvanced` property exposes all 18 ActivityTypes enum values', () => {
		const { ActivityTypes } = require('@microsoft/agents-activity');
		const enumValues = Object.values(ActivityTypes) as string[];
		expect(enumValues.length).toBe(18);

		const trigger = new M365AgentTrigger();
		const activityTypeProp = trigger.description.properties.find(
			(p) => p.name === 'activityTypesAdvanced',
		);
		expect(activityTypeProp).toBeDefined();
		const uiValues = (activityTypeProp?.options as any[]).map((o) => o.value).sort();
		expect(uiValues).toEqual(enumValues.slice().sort());
	});
});

describe('M365AgentTrigger activity-type filter — Simple/Advanced mode split (D1-D)', () => {
	const trigger = new M365AgentTrigger();

	it('Simple mode: accepts a Simple-listed type', async () => {
		// Simple mode active → reads activityTypesSimple (['typing']), accepts 'typing' inbound.
		const ctx = makeTriggerContext(['typing'], 'typing', 'simple');
		const res = await trigger.webhook.call(ctx as any);
		expect(res.workflowData?.[0]?.length ?? 0).toBeGreaterThan(0);
	});

	it('Simple mode: skips a type not in the Simple selection (reads the Simple list when mode=simple)', async () => {
		// Simple selection is ['message'] only → 'typing' must be swallowed.
		const ctx = makeTriggerContext(['message'], 'typing', 'simple');
		const res = await trigger.webhook.call(ctx as any);
		expect(res.workflowData?.[0]?.length ?? 0).toBe(0);
	});

	it('Advanced mode: uses activityTypesAdvanced and ignores the Simple list', async () => {
		// Mode=advanced → Simple list is irrelevant even if set.
		const ctx = makeTriggerContext(['handoff'], 'handoff', 'advanced');
		const res = await trigger.webhook.call(ctx as any);
		expect(res.workflowData?.[0]?.length ?? 0).toBeGreaterThan(0);
	});

	// Metadata assertions for the manifest §6.3 Simple/Advanced pattern.
	it('description exposes an `activityFilterMode` options property with default "simple"', () => {
		const prop = trigger.description.properties.find((p) => p.name === 'activityFilterMode');
		expect(prop).toBeDefined();
		expect(prop?.type).toBe('options');
		expect(prop?.default).toBe('simple');
		const values = (prop?.options as any[]).map((o) => o.value).sort();
		expect(values).toEqual(['advanced', 'simple']);
	});

	it('`activityTypesSimple` exposes exactly the 7 curated types and defaults to [message]', () => {
		const prop = trigger.description.properties.find((p) => p.name === 'activityTypesSimple');
		expect(prop).toBeDefined();
		expect(prop?.default).toEqual(['message']);
		const values = (prop?.options as any[]).map((o) => o.value).sort();
		expect(values).toEqual(
			[
				'conversationUpdate',
				'event',
				'installationUpdate',
				'invoke',
				'message',
				'messageReaction',
				'typing',
			].sort(),
		);
		// Visibility gate: only shown when mode=simple
		expect((prop as any).displayOptions).toEqual({ show: { activityFilterMode: ['simple'] } });
	});

	it('`activityTypesAdvanced` defaults to [message] and is gated on mode=advanced', () => {
		const prop = trigger.description.properties.find((p) => p.name === 'activityTypesAdvanced');
		expect(prop).toBeDefined();
		expect(prop?.default).toEqual(['message']);
		expect((prop as any).displayOptions).toEqual({ show: { activityFilterMode: ['advanced'] } });
	});

	it('old flat `activityTypes` and `invokeNames` properties are removed', () => {
		expect(trigger.description.properties.find((p) => p.name === 'activityTypes')).toBeUndefined();
		expect(trigger.description.properties.find((p) => p.name === 'invokeNames')).toBeUndefined();
	});

	// Counts assertion — the Simple and Advanced options arrays are derived from a
	// single source-of-truth constant in the node file. This guards against drift
	// if the constant is edited without updating the SIMPLE subset filter.
	it('option-list counts are stable: 7 Simple, 18 Advanced', () => {
		const simple = trigger.description.properties.find((p) => p.name === 'activityTypesSimple');
		const advanced = trigger.description.properties.find(
			(p) => p.name === 'activityTypesAdvanced',
		);
		expect((simple?.options as any[]).length).toBe(7);
		expect((advanced?.options as any[]).length).toBe(18);
	});
});
