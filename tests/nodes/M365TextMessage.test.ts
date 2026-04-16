import { describe, it, expect } from 'vitest';
import { M365TextMessage } from '../../nodes/M365TextMessage/M365TextMessage.node';

describe('M365TextMessage description', () => {
	it('has a text property', () => {
		const n = new M365TextMessage();
		const names = n.description.properties.map((p) => p.name);
		expect(names).toContain('text');
	});

	it('is usable as a tool for AI agents', () => {
		const n = new M365TextMessage();
		expect(n.description.usableAsTool).toBe(true);
	});
});
