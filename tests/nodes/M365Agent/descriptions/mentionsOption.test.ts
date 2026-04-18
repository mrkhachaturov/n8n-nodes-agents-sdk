import { describe, it, expect } from 'vitest';
import { mentionsOption } from '../../../../nodes/M365Agent/descriptions/mentionsOption';

describe('mentionsOption property', () => {
	it('is a fixedCollection with multipleValues', () => {
		expect(mentionsOption.type).toBe('fixedCollection');
		expect(mentionsOption.typeOptions?.multipleValues).toBe(true);
	});

	it('has a Values row with type + id + name fields', () => {
		const values = (mentionsOption.options as any[])[0].values;
		const names = values.map((v: any) => v.name);
		expect(names).toEqual(['type', 'id', 'name']);
	});

	it('id field is hidden when type = everyone', () => {
		const values = (mentionsOption.options as any[])[0].values;
		const idField = values.find((v: any) => v.name === 'id');
		expect(idField.displayOptions?.show?.type).toEqual(['user']);
	});

	it('type options include User and Everyone only', () => {
		const values = (mentionsOption.options as any[])[0].values;
		const typeField = values.find((v: any) => v.name === 'type');
		const typeValues = (typeField.options as any[]).map((o) => o.value);
		expect(typeValues.sort()).toEqual(['everyone', 'user']);
	});

	it('multi-sentence field descriptions end with a period (community eslint rule)', () => {
		// The community eslint rule `node-param-description-excess-final-period`
		// forbids a trailing period on single-sentence descriptions but the
		// `node-param-description-missing-final-period` rule requires one when
		// the description is multi-sentence. We assert the project convention:
		// any description that already contains a sentence-ending period inside
		// it must also end with one.
		const values = (mentionsOption.options as any[])[0].values;
		for (const v of values) {
			if (!v.description) continue;
			const inner = v.description.slice(0, -1);
			if (inner.includes('. ')) {
				expect(v.description.endsWith('.')).toBe(true);
			}
		}
	});

	it('property-level description ends with a period (multi-sentence)', () => {
		expect(mentionsOption.description).toBeDefined();
		expect((mentionsOption.description as string).endsWith('.')).toBe(true);
	});
});
