/**
 * 0.5.0 additive-only regression guard.
 *
 * 0.5.0 is the "Tier 3 conversational" milestone: new resources (CardState,
 * InvokeResponse), new routing surfaces (invoke, conversationUpdate,
 * messageReaction), and new envelope fields (invokeName). All additions MUST
 * be additive — no 0.4.0-era field can change shape, no prior resource can
 * disappear, and non-invoke activities must still parse without the new
 * invoke-only fields leaking through.
 *
 * This file is the single-line-of-defense assertion for that contract. If
 * future work shifts from "additive" to "breaking", this test must be
 * deliberately updated — that is the entire point of the regression pin.
 */
import { describe, it, expect } from 'vitest';
import { parseActivity } from '../../shared/envelope';
import { M365Agent } from '../../nodes/M365Agent/M365Agent.node';
import type { Activity } from '@microsoft/agents-activity';
import type { INodePropertyOptions } from 'n8n-workflow';

describe('0.5.0 additive-only regression', () => {
	it('message activity parsing is unchanged (no new required fields)', () => {
		const parsed = parseActivity({
			type: 'message',
			text: 'hello',
			from: { id: 'u', name: 'User' },
		} as unknown as Activity);

		// The pre-0.5.0 ParsedActivity contract — still intact.
		expect(parsed).toBeDefined();
		expect(parsed.type).toBe('message');
		expect(parsed.text).toBe('hello');
		expect(parsed.userId).toBe('u');
		expect(parsed.userName).toBe('User');

		// Conversation/reaction extras must stay absent on plain messages —
		// they're the 0.5.0 additions and must be populated ONLY on their
		// own activity types, never bleed into the message path.
		expect(parsed.membersAdded).toBeUndefined();
		expect(parsed.membersRemoved).toBeUndefined();
		expect(parsed.reactionsAdded).toBeUndefined();
		expect(parsed.reactionsRemoved).toBeUndefined();
	});

	it('envelope invokeName absent on non-invoke activities', () => {
		// parseActivity itself does not populate invokeName — that's the
		// trigger's job on top of the parsed object (see
		// v021-trigger-parity.test.ts for the populated-case coverage).
		// This regression asserts parseActivity has NOT silently started
		// leaking invokeName onto non-invoke activities.
		const parsed = parseActivity({ type: 'message', text: 'x' } as unknown as Activity);
		expect((parsed as { invokeName?: unknown }).invokeName).toBeUndefined();
	});

	it('card-resources routing matrix still handles every card × operation', () => {
		// The pre-0.5.0 ten card resources — NONE may silently disappear.
		// 0.5.0 is allowed to add (cardState, invokeResponse, message); it is
		// not allowed to remove. If a future refactor renames one of these,
		// the test must be updated deliberately alongside a migration note.
		const CARD_RESOURCES = [
			'adaptiveCard',
			'animationCard',
			'audioCard',
			'heroCard',
			'o365ConnectorCard',
			'rawAttachment',
			'receiptCard',
			'signInCard',
			'thumbnailCard',
			'videoCard',
		];

		const node = new M365Agent();
		const resourceProp = node.description.properties.find((p) => p.name === 'resource');
		expect(resourceProp).toBeDefined();
		const values = (resourceProp!.options as INodePropertyOptions[]).map((o) => o.value);

		for (const resource of CARD_RESOURCES) {
			expect(values).toContain(resource);
		}

		// Sanity floor — 0.5.0 ships 13 resources (10 cards + cardState +
		// invokeResponse + message). If the count drops below 13, something
		// was removed rather than added.
		expect(values.length).toBeGreaterThanOrEqual(13);
	});
});
