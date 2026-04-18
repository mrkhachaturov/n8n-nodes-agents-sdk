import { describe, it, expect } from 'vitest';
import { applyMentionsToActivity, buildMentions } from '../../shared/mentions';
import type { Activity } from '@microsoft/agents-activity';

describe('buildMentions', () => {
	it('produces matching text token + entity for a user mention', () => {
		const result = buildMentions([{ type: 'user', id: '29:abc', name: 'Alice' }]);
		expect(result.entities).toEqual([
			{ type: 'mention', mentioned: { id: '29:abc', name: 'Alice' }, text: '<at>Alice</at>' },
		]);
		expect(result.textTokens).toEqual(['<at>Alice</at>']);
	});

	it('uses the Teams magic ID for everyone', () => {
		const result = buildMentions([{ type: 'everyone', name: 'Everyone' }]);
		expect(result.entities[0].mentioned).toEqual({ id: '29:allchannel', name: 'Everyone' });
		expect(result.textTokens).toEqual(['<at>Everyone</at>']);
	});

	it('throws NodeOperationError-compatible on empty name', () => {
		expect(() => buildMentions([{ type: 'user', id: '29:abc', name: '' }])).toThrow(/name/i);
	});

	it('throws on user mention with empty id', () => {
		expect(() => buildMentions([{ type: 'user', id: '', name: 'Alice' }])).toThrow(/id/i);
	});
});

describe('applyMentionsToActivity', () => {
	it('prepends tokens to activity.text and merges entities', () => {
		const activity = { type: 'message', text: 'hello' } as Partial<Activity>;
		const result = applyMentionsToActivity(activity, [
			{ type: 'user', id: '29:abc', name: 'Alice' },
		]);
		expect(result.text).toBe('<at>Alice</at> hello');
		expect(result.entities).toHaveLength(1);
		expect(result.entities?.[0].text).toBe('<at>Alice</at>');
	});

	it('does not double-prepend if token already present in text', () => {
		const activity = { type: 'message', text: '<at>Alice</at> hello' } as Partial<Activity>;
		const result = applyMentionsToActivity(activity, [
			{ type: 'user', id: '29:abc', name: 'Alice' },
		]);
		expect(result.text).toBe('<at>Alice</at> hello');
		expect(result.entities).toHaveLength(1);
	});

	it('preserves existing entities (not-a-mention) on the input activity', () => {
		const activity = {
			type: 'message',
			text: 'hi',
			entities: [{ type: 'clientInfo', locale: 'en-US' }],
		} as Partial<Activity>;
		const result = applyMentionsToActivity(activity, [
			{ type: 'user', id: '29:abc', name: 'Alice' },
		]);
		expect(result.entities).toHaveLength(2);
		expect(result.entities?.some((e) => e.type === 'clientInfo')).toBe(true);
	});

	it('returns a new object, does not mutate input', () => {
		const activity = { type: 'message', text: 'hi' } as Partial<Activity>;
		const before = JSON.stringify(activity);
		applyMentionsToActivity(activity, [{ type: 'user', id: '29:abc', name: 'Alice' }]);
		expect(JSON.stringify(activity)).toBe(before);
	});

	it('empty mentions array returns activity unchanged', () => {
		const activity = { type: 'message', text: 'hi' } as Partial<Activity>;
		const result = applyMentionsToActivity(activity, []);
		expect(result).toEqual(activity);
	});
});
