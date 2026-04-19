import { describe, it, expect } from 'vitest';
import type { IDataObject } from 'n8n-workflow';
import type { Activity } from '@microsoft/agents-activity';

import { applyMessageOptionsToActivity } from '../../../shared/applyMessageOptionsToActivity';

/**
 * G022 — Suggested Actions "To" override.
 *
 * Default behavior is unchanged: when no override is supplied, the third
 * positional arg (`suggestedActionsTo`) auto-populates `activity.suggestedActions.to`
 * from the envelope's clicking user. The override is a proactive-path escape
 * hatch — an explicit list of user IDs that wins over auto-populate.
 *
 * The override lives on the Message `options` object under the name
 * `suggestedActionsToOverride` (peer to `suggestedActions`) to keep the
 * existing fixedCollection shape undisturbed — simpler UX than nesting a
 * string[] field inside the fixedCollection.
 */
describe('Suggested Actions "to" override', () => {
	const baseActivity: Partial<Activity> = { type: 'message', text: 'pick one' };
	const imBackYes = { type: 'imBack', title: 'Yes', value: 'yes' };

	it('auto-populates to from envelope clicker when override is absent (existing behavior)', () => {
		const options: IDataObject = {
			suggestedActions: { values: [imBackYes] },
		};
		const result = applyMessageOptionsToActivity(baseActivity, options, ['inboundUser1']);
		expect(result.suggestedActions?.to).toEqual(['inboundUser1']);
	});

	it('explicit toOverride (non-empty) takes precedence over auto-populate', () => {
		const options: IDataObject = {
			suggestedActions: { values: [imBackYes] },
			suggestedActionsToOverride: ['userA', 'userB'],
		};
		const result = applyMessageOptionsToActivity(baseActivity, options, ['inboundUser1']);
		expect(result.suggestedActions?.to).toEqual(['userA', 'userB']);
	});

	it('override works for proactive path (no envelope clicker)', () => {
		const options: IDataObject = {
			suggestedActions: { values: [imBackYes] },
			suggestedActionsToOverride: ['userA'],
		};
		// empty suggestedActionsTo simulates a proactive send with no inbound user.
		const result = applyMessageOptionsToActivity(baseActivity, options, []);
		expect(result.suggestedActions?.to).toEqual(['userA']);
	});

	it('empty toOverride array means broadcast — no "to" populated even if envelope has a clicker', () => {
		const options: IDataObject = {
			suggestedActions: { values: [imBackYes] },
			suggestedActionsToOverride: [],
		};
		const result = applyMessageOptionsToActivity(baseActivity, options, ['inboundUser1']);
		expect(result.suggestedActions?.to ?? []).toEqual([]);
	});

	it('override ignored when no suggestedActions values are supplied', () => {
		// No chips → no suggestedActions block → override is a no-op.
		const options: IDataObject = {
			suggestedActionsToOverride: ['userA'],
		};
		const result = applyMessageOptionsToActivity(baseActivity, options, ['inboundUser1']);
		expect(result.suggestedActions).toBeUndefined();
	});

	it('non-array toOverride throws with a descriptive message', () => {
		const options: IDataObject = {
			suggestedActions: { values: [imBackYes] },
			// eslint-disable-next-line @typescript-eslint/no-explicit-any
			suggestedActionsToOverride: 'not-an-array' as any,
		};
		expect(() => applyMessageOptionsToActivity(baseActivity, options, [])).toThrow(
			/suggestedActionsToOverride.*array/i,
		);
	});

	it('non-string entries inside toOverride throw', () => {
		const options: IDataObject = {
			suggestedActions: { values: [imBackYes] },
			// eslint-disable-next-line @typescript-eslint/no-explicit-any
			suggestedActionsToOverride: ['valid', 123 as any],
		};
		expect(() => applyMessageOptionsToActivity(baseActivity, options, [])).toThrow(
			/suggestedActionsToOverride/i,
		);
	});
});
