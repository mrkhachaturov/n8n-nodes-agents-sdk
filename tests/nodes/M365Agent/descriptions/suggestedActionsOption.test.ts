import { describe, it, expect } from 'vitest';
import type { INodeProperties, INodePropertyCollection } from 'n8n-workflow';
import {
	suggestedActionsOption,
	suggestedActionsToOverrideOption,
} from '../../../../nodes/M365Agent/descriptions/suggestedActionsOption';

describe('suggestedActionsOption property', () => {
	const valuesRow = (suggestedActionsOption.options as INodePropertyCollection[])[0]
		.values as INodeProperties[];

	it('is a fixedCollection with multipleValues', () => {
		expect(suggestedActionsOption.type).toBe('fixedCollection');
		expect(suggestedActionsOption.typeOptions?.multipleValues).toBe(true);
	});

	it('has Values row with type + title + value + displayText (in display order)', () => {
		const names = valuesRow.map((v) => v.name);
		expect(names).toEqual(['type', 'title', 'value', 'displayText']);
	});

	it('type options are alphabetical by name (manifest community eslint rule)', () => {
		// Case-insensitive sort — matches n8n-nodes-base/node-param-options-type-unsorted-items.
		// The autofixer prefers Title Case for single-word identifiers (Call, Signin) while
		// leaving camelCase ones (imBack, openUrl) alone, so a case-sensitive `.sort()` would
		// split the list; case-insensitive matches what the lint rule actually enforces.
		const typeField = valuesRow.find((v) => v.name === 'type');
		const names = (typeField?.options as Array<{ name: string }> | undefined)?.map((o) => o.name);
		const ci = (a: string, b: string) => a.toLowerCase().localeCompare(b.toLowerCase());
		expect(names).toEqual([...(names ?? [])].sort(ci));
	});

	it('displayText shown only when type = messageBack', () => {
		const dt = valuesRow.find((v) => v.name === 'displayText');
		expect(dt?.displayOptions?.show?.type).toEqual(['messageBack']);
	});

	it('all 11 SDK ActionTypes exposed', () => {
		// Mirrors Microsoft.Agents.Core.Models.ActionTypes (and the JS SDK's
		// ActionTypes enum in agents-activity). The node exposes the full
		// protocol surface; channel support varies but that is not our call.
		const typeField = valuesRow.find((v) => v.name === 'type');
		const typeValues = (typeField?.options as Array<{ value: string }> | undefined)
			?.map((o) => o.value)
			.sort();
		expect(typeValues).toEqual([
			'call',
			'downloadFile',
			'imBack',
			'messageBack',
			'openApp',
			'openUrl',
			'playAudio',
			'playVideo',
			'postBack',
			'showImage',
			'signin',
		]);
	});

	// Repo lint convention (asymmetric period rule, see Task 3 / conversationReference.ts).
	it('field descriptions follow the asymmetric period rule', () => {
		for (const v of valuesRow) {
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

describe('suggestedActionsToOverrideOption property (G022)', () => {
	it('is a multi-string override field with empty-array default', () => {
		expect(suggestedActionsToOverrideOption.name).toBe('suggestedActionsToOverride');
		expect(suggestedActionsToOverrideOption.type).toBe('string');
		expect(suggestedActionsToOverrideOption.typeOptions?.multipleValues).toBe(true);
		expect(suggestedActionsToOverrideOption.default).toEqual([]);
	});

	it('description follows the asymmetric period rule (multi-sentence ends with a period)', () => {
		const desc = suggestedActionsToOverrideOption.description;
		expect(desc).toBeDefined();
		const text = desc as string;
		const isMultiSentence = text.slice(0, -1).includes('. ');
		expect(text.endsWith('.')).toBe(isMultiSentence);
	});
});
