import { describe, it, expect } from 'vitest';
import { M365CardTemplate } from '../../nodes/M365CardTemplate/M365CardTemplate.node';

describe('M365CardTemplate description', () => {
	it('has cardTemplate and bindingData properties', () => {
		const n = new M365CardTemplate();
		const names = n.description.properties.map((p) => p.name);
		expect(names).toContain('cardTemplate');
		expect(names).toContain('bindingData');
	});
});
