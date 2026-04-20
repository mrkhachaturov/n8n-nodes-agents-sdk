import { describe, it, expect } from 'vitest';
import { M365AgentTrigger } from '../../../nodes/M365AgentTrigger/M365AgentTrigger.node';

/**
 * Parsed sub-field E2E coverage for 0.5.0 (Task 6).
 *
 * Task 2 already populates `envelope.parsed.membersAdded` /
 * `envelope.parsed.reactionsAdded` via shared/envelope.ts#parseActivity.
 * This suite pins the Trigger-level behaviour so regressions (e.g. the
 * Trigger dropping `parsed` from the emitted envelope) surface fast.
 *
 * Stub conventions mirror `invoke-name-routing.test.ts` and
 * `activity-type-filter.test.ts`:
 * - `getBodyData`/`getHeaderData` alongside `getRequestObject` (envelope
 *   builder calls the former, signature verification the latter).
 * - `helpers.returnJsonArray` wraps raw items as `[{ json: <item> }]`, so
 *   assertions read `workflowData[0][0].json.parsed.<field>`.
 * - Minimal-but-valid Activity skeleton: `activityToConversationReference`
 *   requires `id`, `serviceUrl`, `channelId`, `from`, `conversation`,
 *   `recipient` — without these the Trigger throws before populating
 *   `parsed`.
 */
function makeCtx(body: Record<string, unknown>) {
	const full: Record<string, unknown> = {
		id: 'a1',
		serviceUrl: 'https://smba/',
		channelId: 'msteams',
		from: { id: '29:u1' },
		conversation: { id: '19:c1', conversationType: 'personal' },
		recipient: { id: '28:b1' },
		...body,
	};
	return {
		getRequestObject: () => ({
			headers: {},
			body: full,
			method: 'POST',
		}),
		getBodyData: () => full,
		getHeaderData: () => ({}),
		getNodeParameter: (name: string) => {
			if (name === 'activityFilterMode') return 'advanced';
			if (name === 'activityTypesSimple') return [];
			if (name === 'activityTypesAdvanced') return [];
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

describe('M365AgentTrigger — parsed sub-fields for 0.5.0', () => {
	const trigger = new M365AgentTrigger();

	it('parsed.membersAdded populated on conversationUpdate', async () => {
		const res = await trigger.webhook.call(
			makeCtx({
				type: 'conversationUpdate',
				membersAdded: [{ id: 'u1', name: 'Alice' }],
			}) as any,
		);
		const envelope = (res.workflowData as any)[0][0].json;
		expect(envelope.parsed.membersAdded).toEqual([{ id: 'u1', name: 'Alice' }]);
	});

	it('parsed.reactionsAdded populated on messageReaction', async () => {
		const res = await trigger.webhook.call(
			makeCtx({
				type: 'messageReaction',
				reactionsAdded: [{ type: 'like', user: { id: 'u1', name: 'Alice' } }],
			}) as any,
		);
		const envelope = (res.workflowData as any)[0][0].json;
		expect(envelope.parsed.reactionsAdded).toEqual([
			{ type: 'like', user: { id: 'u1', name: 'Alice' } },
		]);
	});

	it('parsed sub-fields undefined on plain message', async () => {
		const res = await trigger.webhook.call(
			makeCtx({ type: 'message', text: 'hello' }) as any,
		);
		const envelope = (res.workflowData as any)[0][0].json;
		expect(envelope.parsed.membersAdded).toBeUndefined();
		expect(envelope.parsed.reactionsAdded).toBeUndefined();
	});
});
