import { describe, it, expect } from 'vitest';
import type { IDataObject } from 'n8n-workflow';
import type { Activity } from '@microsoft/agents-activity';

import { applyMessageOptionsToActivity } from '../../../shared/applyMessageOptionsToActivity';

/**
 * Suggested Actions "To" override — now nested inside the Suggested Actions
 * collection per spec §4 line 113 and §8 line 220 (audit A3/D2-A). The chip
 * entries live at `options.suggestedActions.chips.values` and the override
 * lives at `options.suggestedActions.to` (string[]).
 *
 * Default behavior is unchanged: when no override is supplied, the third
 * positional arg (`suggestedActionsTo`) auto-populates `activity.suggestedActions.to`
 * from the envelope's clicking user. The override is a proactive-path escape
 * hatch — an explicit list of user IDs that wins over auto-populate.
 */
describe('Suggested Actions "to" override', () => {
	const baseActivity: Partial<Activity> = { type: 'message', text: 'pick one' };
	const imBackYes = { type: 'imBack', title: 'Yes', value: 'yes' };

	it('auto-populates to from envelope clicker when override is absent (existing behavior)', () => {
		const options: IDataObject = {
			suggestedActions: { chips: { values: [imBackYes] } },
		};
		const result = applyMessageOptionsToActivity(baseActivity, options, ['inboundUser1']);
		expect(result.suggestedActions?.to).toEqual(['inboundUser1']);
	});

	it('explicit to override (non-empty) takes precedence over auto-populate', () => {
		const options: IDataObject = {
			suggestedActions: {
				chips: { values: [imBackYes] },
				to: ['userA', 'userB'],
			},
		};
		const result = applyMessageOptionsToActivity(baseActivity, options, ['inboundUser1']);
		expect(result.suggestedActions?.to).toEqual(['userA', 'userB']);
	});

	it('override works for proactive path (no envelope clicker)', () => {
		const options: IDataObject = {
			suggestedActions: {
				chips: { values: [imBackYes] },
				to: ['userA'],
			},
		};
		// empty suggestedActionsTo simulates a proactive send with no inbound user.
		const result = applyMessageOptionsToActivity(baseActivity, options, []);
		expect(result.suggestedActions?.to).toEqual(['userA']);
	});

	it('empty to override array means broadcast — no "to" populated even if envelope has a clicker', () => {
		const options: IDataObject = {
			suggestedActions: {
				chips: { values: [imBackYes] },
				to: [],
			},
		};
		const result = applyMessageOptionsToActivity(baseActivity, options, ['inboundUser1']);
		expect(result.suggestedActions?.to ?? []).toEqual([]);
	});

	it('override ignored when no suggestedActions chips are supplied', () => {
		// No chips → no suggestedActions block → override is a no-op.
		const options: IDataObject = {
			suggestedActions: { to: ['userA'] },
		};
		const result = applyMessageOptionsToActivity(baseActivity, options, ['inboundUser1']);
		expect(result.suggestedActions).toBeUndefined();
	});

	it('non-array to override throws with a descriptive message', () => {
		const options: IDataObject = {
			suggestedActions: {
				chips: { values: [imBackYes] },
				// eslint-disable-next-line @typescript-eslint/no-explicit-any
				to: 'not-an-array' as any,
			},
		};
		expect(() => applyMessageOptionsToActivity(baseActivity, options, [])).toThrow(
			/suggestedActions.*to.*array/i,
		);
	});

	it('non-string entries inside to override throw', () => {
		const options: IDataObject = {
			suggestedActions: {
				chips: { values: [imBackYes] },
				// eslint-disable-next-line @typescript-eslint/no-explicit-any
				to: ['valid', 123 as any],
			},
		};
		expect(() => applyMessageOptionsToActivity(baseActivity, options, [])).toThrow(
			/suggestedActions.*to/i,
		);
	});
});
