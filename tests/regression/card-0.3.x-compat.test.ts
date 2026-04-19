import { describe, it, expect, vi } from 'vitest';
import { M365Agent } from '../../nodes/M365Agent/M365Agent.node';
import { makeExecuteContext, makeCredentials } from '../helpers/makeContext';
import type { IExecuteFunctions } from 'n8n-workflow';

vi.mock('../../shared/auth/router', () => ({
	acquireOutboundToken: vi.fn().mockResolvedValue({ authorizationHeader: 'Bearer fake' }),
}));
vi.mock('../../shared/botConnector', () => ({
	createConnectorFromBearer: vi.fn().mockReturnValue({
		client: { sendToConversation: vi.fn(), updateActivity: vi.fn() },
		axios: {},
		token: 'fake',
		baseURL: 'https://smba/',
	}),
	replyInThread: vi.fn(),
}));

const ENVELOPE = {
	conversationReference: {
		serviceUrl: 'https://smba/',
		conversation: { id: '19:c@thread.tacv2' },
		channelId: 'msteams',
	},
};

describe('0.3.x compatibility — hard break', () => {
	it("a workflow saved against 0.3.x using resource: 'card' now fails loudly with the full preserved message", async () => {
		const node = new M365Agent();
		const ctx = makeExecuteContext({
			inputItems: [ENVELOPE],
			credentials: makeCredentials(),
			parameters: {
				resource: 'card',
				operation: 'send',
				authKind: 'classicBot',
				conversationSource: 'envelope',
				cardTemplate: '{"type":"AdaptiveCard","version":"1.4","body":[]}',
				bindingData: {},
				options: {},
			},
		});
		// The spec-preserved default-path message has two load-bearing parts:
		//   (1) the `Unexpected resource "<name>"` prefix (6c from spec §5.1)
		//   (2) the invokeResponse single-HTTP-request invariant reminder.
		// Both must be asserted — a test that only matched the prefix would
		// silently accept a refactor that stripped the invariant text.
		const promise = node.execute.call(ctx as unknown as IExecuteFunctions);
		await expect(promise).rejects.toThrow(/Unexpected resource "card"/);
		await expect(promise).rejects.toThrow(
			/Invoke Response must be the only operation in this node execution/,
		);
		await expect(promise).rejects.toThrow(
			/responds to a single HTTP request and cannot be combined with other resources/,
		);
	});
});
