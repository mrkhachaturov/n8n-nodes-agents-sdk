import { describe, it, expect } from 'vitest';
import { M365TextMessage } from '../../nodes/M365TextMessage/M365TextMessage.node';
import { makeExecuteContext } from '../helpers/makeContext';

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

// ---------------------------------------------------------------------------
// execute() behaviour tests
// ---------------------------------------------------------------------------

describe('M365TextMessage execute()', () => {
	it('preserves conversationReference when present and replaces activity', async () => {
		const conversationReference = {
			serviceUrl: 'https://smba.trafficmanager.net/amer/',
			conversation: { id: 'conv-001' },
			activityId: 'act-001',
			channelId: 'msteams',
		};
		const oldActivity = { type: 'message', text: 'old text' };

		const ctx = makeExecuteContext({
			inputItems: [{ conversationReference, activity: oldActivity }],
			parameters: { text: 'Hello from n8n!' },
		});

		const node = new M365TextMessage();
		const [output] = await node.execute.call(ctx as never);

		expect(output).toHaveLength(1);
		const json = output[0].json as Record<string, unknown>;
		// conversationReference must be the original object, untouched
		expect(json.conversationReference).toEqual(conversationReference);
		// activity must be the new one built by the node
		expect(json.activity).toMatchObject({ type: 'message', text: 'Hello from n8n!' });
	});

	it('emits activity-only when conversationReference is absent', async () => {
		const ctx = makeExecuteContext({
			inputItems: [{ activity: { type: 'message', text: 'stub' } }],
			parameters: { text: 'Proactive message' },
		});

		const node = new M365TextMessage();
		const [output] = await node.execute.call(ctx as never);

		expect(output).toHaveLength(1);
		const json = output[0].json as Record<string, unknown>;
		expect(json.conversationReference).toBeUndefined();
		expect(json.activity).toMatchObject({ type: 'message', text: 'Proactive message' });
	});
});
