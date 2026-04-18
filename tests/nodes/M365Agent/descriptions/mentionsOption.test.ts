import { describe, it, expect } from 'vitest';
import type { INodeProperties, INodePropertyCollection } from 'n8n-workflow';
import { mentionsOption } from '../../../../nodes/M365Agent/descriptions/mentionsOption';

describe('mentionsOption property', () => {
	const valuesRow = (mentionsOption.options as INodePropertyCollection[])[0]
		.values as INodeProperties[];

	it('is a fixedCollection with multipleValues', () => {
		expect(mentionsOption.type).toBe('fixedCollection');
		expect(mentionsOption.typeOptions?.multipleValues).toBe(true);
	});

	it('has a Values row with type + id + name fields', () => {
		const names = valuesRow.map((v) => v.name);
		expect(names).toEqual(['type', 'id', 'name']);
	});

	it('id field is hidden when type = everyone', () => {
		const idField = valuesRow.find((v) => v.name === 'id');
		expect(idField?.displayOptions?.show?.type).toEqual(['user']);
	});

	it('type options include User and Everyone only', () => {
		const typeField = valuesRow.find((v) => v.name === 'type');
		const typeValues = (typeField?.options as Array<{ value: string }> | undefined)?.map(
			(o) => o.value,
		);
		expect(typeValues?.sort()).toEqual(['everyone', 'user']);
	});

	it('all field descriptions follow the asymmetric period rule', () => {
		// The community eslint rule `node-param-description-excess-final-period`
		// forbids a trailing period on single-sentence descriptions but the
		// `node-param-description-missing-final-period` rule requires one when
		// the description is multi-sentence. We assert the project convention:
		// any description that already contains a sentence-ending period inside
		// it must also end with one.
		let assertions = 0;
		for (const v of valuesRow) {
			if (!v.description) continue;
			const isMultiSentence = v.description.slice(0, -1).includes('. ');
			expect(v.description.endsWith('.')).toBe(isMultiSentence);
			assertions++;
		}
		expect(assertions).toBeGreaterThan(0);
	});

	it('property-level description ends with a period (multi-sentence)', () => {
		expect(mentionsOption.description).toBeDefined();
		expect((mentionsOption.description as string).endsWith('.')).toBe(true);
	});
});
