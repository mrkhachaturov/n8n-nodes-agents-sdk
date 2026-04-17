import { describe, it, expect } from 'vitest';
import * as card from '../../../nodes/M365Agent/actions/card';

describe('card resource description', () => {
	it('operation selector is scoped to card resource, alphabetical, noDataExpression', () => {
		const op = card.description.find((p) => p.name === 'operation');
		expect(op?.noDataExpression).toBe(true);
		const shown = (op?.displayOptions?.show as Record<string, string[]> | undefined)?.resource;
		expect(shown).toEqual(['card']);
		const values = (op?.options as { value: string }[] | undefined)?.map((o) => o.value);
		expect(values).toEqual(['send']);
		expect(op?.default).toBe('send');
	});

	it('cardTemplate field shown only for card/send', () => {
		const tmpl = card.description.find((p) => p.name === 'cardTemplate');
		const shown = tmpl?.displayOptions?.show as Record<string, string[]> | undefined;
		expect(shown?.resource).toEqual(['card']);
		expect(shown?.operation).toEqual(['send']);
	});
});
