import { describe, it, expect } from 'vitest';
import type { Activity } from '@microsoft/agents-activity';
import type { IDataObject } from 'n8n-workflow';

import { applyMessageOptionsToActivity } from '../../shared/applyMessageOptionsToActivity';

describe('applyMessageOptionsToActivity', () => {
	const baseActivity: Partial<Activity> = { type: 'message', text: 'hello' };

	it('returns activity unchanged when options is empty {}', () => {
		const options: IDataObject = {};
		const result = applyMessageOptionsToActivity(baseActivity, options);
		expect(result).toEqual(baseActivity);
	});

	it('applies only mentions when only options.mentions.values is non-empty', () => {
		const options: IDataObject = {
			mentions: { values: [{ type: 'user', id: '29:abc', name: 'Alice' }] },
		};
		const result = applyMessageOptionsToActivity(baseActivity, options);
		expect(result.text).toBe('<at>Alice</at> hello');
		expect(result.entities).toHaveLength(1);
		expect(result.entities?.[0].type).toBe('mention');
		// suggestedActions should NOT be set
		expect(result.suggestedActions).toBeUndefined();
	});

	it('applies only suggestedActions when only options.suggestedActions.values is non-empty', () => {
		const options: IDataObject = {
			suggestedActions: {
				values: [{ type: 'imBack', title: 'Yes', value: 'yes' }],
			},
		};
		const result = applyMessageOptionsToActivity(baseActivity, options);
		// text untouched, no entities added
		expect(result.text).toBe('hello');
		expect(result.entities).toBeUndefined();
		expect(result.suggestedActions?.actions).toHaveLength(1);
		expect(result.suggestedActions?.actions?.[0]).toMatchObject({
			type: 'imBack',
			title: 'Yes',
			value: 'yes',
		});
	});

	it('applies both mentions then suggestedActions when both present', () => {
		const options: IDataObject = {
			mentions: { values: [{ type: 'user', id: '29:abc', name: 'Alice' }] },
			suggestedActions: {
				values: [{ type: 'imBack', title: 'Yes', value: 'yes' }],
			},
		};
		const result = applyMessageOptionsToActivity(baseActivity, options);
		// mention applied to text + entities
		expect(result.text).toBe('<at>Alice</at> hello');
		expect(result.entities).toHaveLength(1);
		// suggested actions also present
		expect(result.suggestedActions?.actions).toHaveLength(1);
		expect(result.suggestedActions?.actions?.[0].title).toBe('Yes');
	});

	it('propagates Error from buildMentions on invalid mention shape', () => {
		const options: IDataObject = {
			mentions: { values: [{ type: 'user', id: '', name: 'X' }] },
		};
		expect(() => applyMessageOptionsToActivity(baseActivity, options)).toThrow(Error);
		expect(() => applyMessageOptionsToActivity(baseActivity, options)).toThrow(/id/i);
	});

	it('propagates Error from buildSuggestedActions on invalid action shape', () => {
		const options: IDataObject = {
			suggestedActions: {
				values: [{ type: 'imBack', title: '', value: 'x' }],
			},
		};
		expect(() => applyMessageOptionsToActivity(baseActivity, options)).toThrow(Error);
		expect(() => applyMessageOptionsToActivity(baseActivity, options)).toThrow(/title/i);
	});
});
