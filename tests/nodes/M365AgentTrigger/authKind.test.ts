import { describe, it, expect } from 'vitest';
import { M365AgentTrigger } from '../../../nodes/M365AgentTrigger/M365AgentTrigger.node';

describe('M365AgentTrigger — authKind param + credential gating', () => {
	const t = new M365AgentTrigger();
	const props = t.description.properties!;

	it('exposes authKind as first parameter', () => {
		expect(props[0].name).toBe('authKind');
		expect(props[0].type).toBe('options');
		expect(props[0].noDataExpression).toBe(true);
		expect(props[0].default).toBe('classicBot');
	});

	it('declares both credential types gated by authKind', () => {
		const creds = t.description.credentials!;
		const names = creds.map((c) => c.name);
		expect(names).toEqual(expect.arrayContaining(['m365AgentApi', 'm365Agent365Api']));
		const classic = creds.find((c) => c.name === 'm365AgentApi')!;
		expect(classic.displayOptions?.show).toEqual({ authKind: ['classicBot'] });
		const agent365 = creds.find((c) => c.name === 'm365Agent365Api')!;
		expect(agent365.displayOptions?.show).toEqual({ authKind: ['agent365'] });
	});

	it('exposes agent365CredentialTest method', () => {
		expect(t.methods?.credentialTest?.agent365CredentialTest).toBeTypeOf('function');
	});
});
