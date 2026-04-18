import { describe, it, expect } from 'vitest';
import { M365Agent } from '../../../nodes/M365Agent/M365Agent.node';

describe('M365Agent — authKind param + credential gating', () => {
	const n = new M365Agent();
	const props = n.description.properties!;

	it('first property is authKind', () => {
		expect(props[0].name).toBe('authKind');
		expect(props[0].default).toBe('classicBot');
	});

	it('credentials declare both types with displayOptions gating', () => {
		const creds = n.description.credentials!;
		expect(creds).toHaveLength(2);
		expect(creds.find((c) => c.name === 'm365AgentApi')?.displayOptions?.show).toEqual({
			authKind: ['classicBot'],
		});
		expect(creds.find((c) => c.name === 'm365Agent365Api')?.displayOptions?.show).toEqual({
			authKind: ['agent365'],
		});
	});
});
