import { describe, it, expect } from 'vitest';
import { applyRawActivityOverride, rawActivityOverrideField } from '../../shared/rawActivityOverride';

describe('applyRawActivityOverride', () => {
	const constructed = { type: 'message', text: 'constructed body' };

	it('returns the constructed activity when override is empty', () => {
		expect(applyRawActivityOverride(constructed, '')).toEqual(constructed);
		expect(applyRawActivityOverride(constructed, undefined)).toEqual(constructed);
		expect(applyRawActivityOverride(constructed, null as any)).toEqual(constructed);
	});

	it('replaces the activity body with parsed override JSON', () => {
		const override = '{"type":"message","text":"overridden","channelData":{"teamsExtra":true}}';
		const result = applyRawActivityOverride(constructed, override);
		expect(result).toEqual({ type: 'message', text: 'overridden', channelData: { teamsExtra: true } });
	});

	it('throws a typed error with field name on invalid JSON', () => {
		expect(() => applyRawActivityOverride(constructed, '{not json'))
			.toThrow(/rawActivityOverride.*invalid JSON/i);
	});

	it('throws when override is not an object', () => {
		expect(() => applyRawActivityOverride(constructed, '"just a string"'))
			.toThrow(/rawActivityOverride.*expected JSON object/i);
	});

	it('exports a reusable INodeProperties field definition', () => {
		expect(rawActivityOverrideField.name).toBe('rawActivityOverride');
		expect(rawActivityOverrideField.type).toBe('string');
		expect(rawActivityOverrideField.default).toBe('');
	});
});
