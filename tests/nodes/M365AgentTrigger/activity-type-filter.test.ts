import { describe, it, expect } from 'vitest';
import { M365AgentTrigger } from '../../../nodes/M365AgentTrigger/M365AgentTrigger.node';

/**
 * The filter is implemented in webhook(). We stub IWebhookFunctions
 * and exercise each of the 18 ActivityTypes enum values.
 */
const ALL_18_TYPES = [
	'message', 'contactRelationUpdate', 'conversationUpdate', 'typing',
	'endOfConversation', 'event', 'invoke', 'invokeResponse', 'deleteUserData',
	'messageUpdate', 'messageDelete', 'installationUpdate', 'messageReaction',
	'suggestion', 'trace', 'handoff', 'command', 'commandResult',
];

function makeTriggerContext(configuredFilter: string[], inboundType: string) {
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
			if (name === 'activityTypes') return configuredFilter;
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

describe('M365AgentTrigger activity-type filter — all 18 types', () => {
	const trigger = new M365AgentTrigger();

	for (const type of ALL_18_TYPES) {
		it(`accepts ${type} when filter includes it`, async () => {
			const ctx = makeTriggerContext([type], type);
			const res = await trigger.webhook.call(ctx as any);
			expect(res.workflowData?.[0]?.length ?? 0).toBeGreaterThan(0);
		});

		it(`skips ${type} when filter does not include it`, async () => {
			const otherType = type === 'message' ? 'typing' : 'message';
			const ctx = makeTriggerContext([otherType], type);
			const res = await trigger.webhook.call(ctx as any);
			expect(res.workflowData?.[0]?.length ?? 0).toBe(0);
		});
	}

	it('accepts any type when filter is empty (default = no filter)', async () => {
		const ctx = makeTriggerContext([], 'handoff');
		const res = await trigger.webhook.call(ctx as any);
		expect(res.workflowData?.[0]?.length ?? 0).toBeGreaterThan(0);
	});

	// Metadata assertion — protects against the UI drifting from the authoritative enum.
	it('Trigger `activityTypes` property exposes all 18 ActivityTypes enum values', () => {
		const { ActivityTypes } = require('@microsoft/agents-activity');
		const enumValues = Object.values(ActivityTypes) as string[];
		expect(enumValues.length).toBe(18);

		const trigger = new M365AgentTrigger();
		const activityTypeProp = trigger.description.properties.find((p) => p.name === 'activityTypes');
		expect(activityTypeProp).toBeDefined();
		const uiValues = (activityTypeProp?.options as any[]).map((o) => o.value).sort();
		expect(uiValues).toEqual(enumValues.slice().sort());
	});
});
