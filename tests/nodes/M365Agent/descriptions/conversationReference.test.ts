import { describe, it, expect } from 'vitest';
import { conversationReferenceProperties } from '../../../../nodes/M365Agent/descriptions/conversationReference';

describe('conversationReference properties', () => {
	it('exposes a Conversation Source mode with Simple/Advanced options (as-written order)', () => {
		const mode = conversationReferenceProperties.find((p) => p.name === 'conversationSource');
		expect(mode?.type).toBe('options');
		expect(mode?.noDataExpression).toBe(true);
		const names = (mode?.options as { name: string }[] | undefined)?.map((o) => o.name);
		// Manifest §14 requires alphabetical by name; assert as written.
		expect(names).toEqual(['From Envelope', 'Specify Manually']);
		expect(mode?.default).toBe('envelope');
	});

	it('every non-mode property gates on conversationSource=manual', () => {
		const nonMode = conversationReferenceProperties.filter((p) => p.name !== 'conversationSource');
		// Spread every item's displayOptions.show.conversationSource; all must be ['manual'].
		for (const p of nonMode) {
			const shown = (p.displayOptions?.show as Record<string, string[]> | undefined)
				?.conversationSource;
			expect(shown).toEqual(['manual']);
		}
	});

	it('Advanced mode exposes exactly serviceUrl, conversationId, channelId, activityId (in order)', () => {
		const manualFieldNames = conversationReferenceProperties
			.filter((p) => p.name !== 'conversationSource')
			.map((p) => p.name);
		expect(manualFieldNames).toEqual(['serviceUrl', 'conversationId', 'channelId', 'activityId']);
	});

	it('serviceUrl and conversationId are required; activityId is optional', () => {
		const svc = conversationReferenceProperties.find((p) => p.name === 'serviceUrl');
		const conv = conversationReferenceProperties.find((p) => p.name === 'conversationId');
		const act = conversationReferenceProperties.find((p) => p.name === 'activityId');
		expect(svc?.required).toBe(true);
		expect(conv?.required).toBe(true);
		expect(act?.required).toBeFalsy();
	});
});
