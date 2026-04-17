import { describe, it, expect } from 'vitest';
import { M365Agent } from '../../../nodes/M365Agent/M365Agent.node';

describe('M365Agent node', () => {
	it('declares displayName "M365 Agent"', () => {
		expect(new M365Agent().description.displayName).toBe('M365 Agent');
	});
	it('uses m365AgentApi credential', () => {
		expect(new M365Agent().description.credentials?.[0]?.name).toBe('m365AgentApi');
	});
	it('does not expose usableAsTool (manifest §10.7 — side-effecting node)', () => {
		expect(new M365Agent().description.usableAsTool).toBeFalsy();
	});
});
