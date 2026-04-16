import { describe, it, expect } from 'vitest';
import { M365SendActivity } from '../../nodes/M365SendActivity/M365SendActivity.node';

describe('M365SendActivity description', () => {
	it('has all five operations', () => {
		const n = new M365SendActivity();
		const op = n.description.properties.find((p) => p.name === 'operation');
		const values = (op?.options as { value: string }[] | undefined)?.map((o) => o.value).sort();
		expect(values).toEqual(['delete', 'proactive', 'reply', 'replyInThread', 'update']);
	});

	it('exposes parentActivityId for replyInThread only', () => {
		const n = new M365SendActivity();
		const p = n.description.properties.find((x) => x.name === 'parentActivityId');
		const shown = (p?.displayOptions?.show as Record<string, unknown> | undefined)?.operation;
		expect(shown).toEqual(['replyInThread']);
	});

	it('binds to m365AgentApi credential', () => {
		const n = new M365SendActivity();
		const credName = (n.description.credentials ?? [])[0]?.name;
		expect(credName).toBe('m365AgentApi');
	});
});
