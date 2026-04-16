import { describe, it, expect } from 'vitest';
import { M365AgentTrigger } from '../../nodes/M365AgentTrigger/M365AgentTrigger.node';

describe('M365AgentTrigger description', () => {
	it('registers two webhooks: POST messages + GET health', () => {
		const node = new M365AgentTrigger();
		const methods = (node.description.webhooks ?? []).map((w) => w.httpMethod).sort();
		expect(methods).toEqual(['GET', 'POST']);
	});

	it('binds to m365AgentApi credential', () => {
		const node = new M365AgentTrigger();
		const credName = (node.description.credentials ?? [])[0]?.name;
		expect(credName).toBe('m365AgentApi');
	});

	it('exposes activityTypes and channelFilter properties', () => {
		const node = new M365AgentTrigger();
		const names = node.description.properties.map((p) => p.name);
		expect(names).toContain('activityTypes');
		expect(names).toContain('channelFilter');
	});

	it('does NOT expose a raw output format (envelope contract is fixed)', () => {
		const node = new M365AgentTrigger();
		const names = node.description.properties.map((p) => p.name);
		expect(names).not.toContain('outputFormat');
	});
});
