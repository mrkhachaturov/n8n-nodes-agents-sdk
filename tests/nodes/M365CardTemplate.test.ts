import { describe, it, expect } from 'vitest';
import { M365CardTemplate } from '../../nodes/M365CardTemplate/M365CardTemplate.node';
import { makeExecuteContext } from '../helpers/makeContext';

describe('M365CardTemplate description', () => {
	it('has cardTemplate and bindingData properties', () => {
		const n = new M365CardTemplate();
		const names = n.description.properties.map((p) => p.name);
		expect(names).toContain('cardTemplate');
		expect(names).toContain('bindingData');
	});
});

// ---------------------------------------------------------------------------
// execute() behaviour tests
// ---------------------------------------------------------------------------

describe('M365CardTemplate execute()', () => {
	it('expands ${field} placeholders from bindingData into the card', async () => {
		const template = {
			type: 'AdaptiveCard',
			'$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
			version: '1.4',
			body: [{ type: 'TextBlock', text: '${title}' }],
		};
		const binding = { title: 'Hello from binding' };
		const conversationReference = {
			serviceUrl: 'https://smba.trafficmanager.net/amer/',
			conversation: { id: 'conv-001' },
			channelId: 'msteams',
		};

		const ctx = makeExecuteContext({
			inputItems: [{ conversationReference }],
			parameters: {
				cardTemplate: JSON.stringify(template),
				bindingData: JSON.stringify(binding),
			},
		});

		const node = new M365CardTemplate();
		const [output] = await node.execute.call(ctx as never);

		expect(output).toHaveLength(1);
		const json = output[0].json as Record<string, unknown>;
		const activity = json.activity as Record<string, unknown>;
		expect(activity.type).toBe('message');

		const attachments = activity.attachments as Array<Record<string, unknown>>;
		expect(attachments).toHaveLength(1);
		expect(attachments[0].contentType).toBe('application/vnd.microsoft.card.adaptive');

		const content = attachments[0].content as Record<string, unknown>;
		const body = content.body as Array<Record<string, unknown>>;
		expect(body[0].text).toBe('Hello from binding');
	});

	it('throws NodeOperationError when template expansion fails', async () => {
		// Provide a truncated (invalid) JSON string for cardTemplate — JSON.parse inside
		// the node will throw, which gets caught and re-thrown as NodeOperationError.
		const ctx = makeExecuteContext({
			inputItems: [{}],
			parameters: {
				cardTemplate: '{"type":"AdaptiveCard","body":[{"type":"TextBlock","text":"${title}"',  // truncated — invalid JSON
				bindingData: '{}',
			},
		});

		const node = new M365CardTemplate();
		await expect(node.execute.call(ctx as never)).rejects.toThrow();
	});
});
