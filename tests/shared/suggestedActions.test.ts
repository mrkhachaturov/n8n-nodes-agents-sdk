import { describe, it, expect } from 'vitest';
import {
	applySuggestedActionsToActivity,
	buildSuggestedActions,
} from '../../shared/suggestedActions';
import type { Activity } from '@microsoft/agents-activity';

describe('buildSuggestedActions', () => {
	it('maps imBack action input to CardAction', () => {
		const result = buildSuggestedActions([{ type: 'imBack', title: 'Yes', value: 'yes' }]);
		expect(result).toEqual({
			to: [],
			actions: [{ type: 'imBack', title: 'Yes', value: 'yes' }],
		});
	});

	it('includes displayText for messageBack', () => {
		const result = buildSuggestedActions([
			{ type: 'messageBack', title: 'Approve', value: 'approved', displayText: '👍 Approved' },
		]);
		expect(result.actions[0]).toEqual({
			type: 'messageBack',
			title: 'Approve',
			value: 'approved',
			displayText: '👍 Approved',
		});
	});

	it('supports openUrl with a URL value', () => {
		const result = buildSuggestedActions([
			{ type: 'openUrl', title: 'Docs', value: 'https://example.com/docs' },
		]);
		expect(result.actions[0]).toEqual({
			type: 'openUrl',
			title: 'Docs',
			value: 'https://example.com/docs',
		});
	});

	it('throws on empty title', () => {
		expect(() => buildSuggestedActions([{ type: 'imBack', title: '', value: 'x' }])).toThrow(
			/title/i,
		);
	});

	it('throws on unknown type', () => {
		// Force an invalid type to exercise the validator.
		expect(() =>
			buildSuggestedActions([{ type: 'bogus' as any, title: 't', value: 'v' }]),
		).toThrow(/type/i);
	});

	it('throws on empty value (all action types)', () => {
		expect(() => buildSuggestedActions([{ type: 'imBack', title: 't', value: '' }])).toThrow(
			/value/i,
		);
	});

	it('drops displayText on non-messageBack action types', () => {
		const result = buildSuggestedActions([
			// displayText supplied but should be silently dropped for imBack
			{ type: 'imBack', title: 't', value: 'v', displayText: 'should-be-dropped' },
		]);
		expect(result.actions[0]).toEqual({ type: 'imBack', title: 't', value: 'v' });
		expect(result.actions[0]).not.toHaveProperty('displayText');
	});

	it('messageBack without displayText omits the field (does not set undefined)', () => {
		const result = buildSuggestedActions([
			{ type: 'messageBack', title: 't', value: 'v' },
		]);
		expect(result.actions[0]).toEqual({ type: 'messageBack', title: 't', value: 'v' });
		expect(result.actions[0]).not.toHaveProperty('displayText');
	});
});

describe('applySuggestedActionsToActivity', () => {
	it('sets activity.suggestedActions from inputs', () => {
		const activity = { type: 'message', text: 'pick one' } as Partial<Activity>;
		const result = applySuggestedActionsToActivity(activity, [
			{ type: 'imBack', title: 'A', value: 'a' },
			{ type: 'imBack', title: 'B', value: 'b' },
		]);
		expect(result.suggestedActions?.actions).toHaveLength(2);
		expect(result.suggestedActions?.to).toEqual([]);
	});

	it('empty inputs returns activity unchanged', () => {
		const activity = { type: 'message', text: 'hi' } as Partial<Activity>;
		const result = applySuggestedActionsToActivity(activity, []);
		expect(result).toEqual(activity);
	});

	it('does not mutate the input activity', () => {
		const activity = { type: 'message', text: 'hi' } as Partial<Activity>;
		const before = JSON.stringify(activity);
		applySuggestedActionsToActivity(activity, [{ type: 'imBack', title: 'a', value: 'a' }]);
		expect(JSON.stringify(activity)).toBe(before);
	});
});
