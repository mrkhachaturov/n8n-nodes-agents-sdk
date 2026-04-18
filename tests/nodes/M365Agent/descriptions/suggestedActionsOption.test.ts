import { describe, it, expect } from 'vitest';
import { suggestedActionsOption } from '../../../../nodes/M365Agent/descriptions/suggestedActionsOption';

describe('suggestedActionsOption property', () => {
	it('is a fixedCollection with multipleValues', () => {
		expect(suggestedActionsOption.type).toBe('fixedCollection');
		expect(suggestedActionsOption.typeOptions?.multipleValues).toBe(true);
	});

	it('has Values row with type + title + value + displayText', () => {
		const values = (suggestedActionsOption.options as any[])[0].values;
		const names = values.map((v: any) => v.name);
		expect(names.sort()).toEqual(['displayText', 'title', 'type', 'value']);
	});

	it('type options are alphabetical by name (manifest community eslint rule)', () => {
		const values = (suggestedActionsOption.options as any[])[0].values;
		const typeField = values.find((v: any) => v.name === 'type');
		const names = (typeField.options as any[]).map((o) => o.name);
		expect(names).toEqual([...names].sort());
	});

	it('displayText shown only when type = messageBack', () => {
		const values = (suggestedActionsOption.options as any[])[0].values;
		const dt = values.find((v: any) => v.name === 'displayText');
		expect(dt.displayOptions?.show?.type).toEqual(['messageBack']);
	});

	it('four action types exposed (M1 scope)', () => {
		const values = (suggestedActionsOption.options as any[])[0].values;
		const typeField = values.find((v: any) => v.name === 'type');
		const typeValues = (typeField.options as any[]).map((o) => o.value).sort();
		expect(typeValues).toEqual(['imBack', 'messageBack', 'openUrl', 'postBack']);
	});

	// Repo lint convention (asymmetric period rule, see Task 3 / conversationReference.ts).
	it('field descriptions follow the asymmetric period rule', () => {
		const values = (suggestedActionsOption.options as any[])[0].values;
		for (const v of values) {
			if (!v.description) continue;
			const isMultiSentence = v.description.slice(0, -1).includes('. ');
			expect(v.description.endsWith('.')).toBe(isMultiSentence);
		}
	});

	it('property-level description ends with a period (multi-sentence)', () => {
		expect(suggestedActionsOption.description).toBeDefined();
		expect((suggestedActionsOption.description as string).endsWith('.')).toBe(true);
	});
});
