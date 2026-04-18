import { describe, it, expect, vi } from 'vitest';
import { M365Agent } from '../../../nodes/M365Agent/M365Agent.node';
import { makeExecuteContext, makeCredentials } from '../../helpers/makeContext';
import type { IExecuteFunctions } from 'n8n-workflow';

vi.mock('../../../shared/auth/router', () => ({
	acquireOutboundToken: vi.fn().mockResolvedValue({ authorizationHeader: 'Bearer fake-token' }),
}));

vi.mock('../../../shared/botConnector', () => ({
	createConnectorFromBearer: vi.fn(),
	replyInThread: vi.fn(),
}));

// Drive the connector mock from the regression-guard test below.
import { createConnectorFromBearer as mockCreateConnectorFromBearer } from '../../../shared/botConnector';

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

// Defence-in-depth: the router's `default` branches are only reachable if a new
// operation is wired into the description without a matching case. These tests
// lock the `NodeOperationError` contract so a silent TypeError can't slip in.
describe('M365Agent router — unknown operation defaults', () => {
	it('throws NodeOperationError for an unknown message operation', async () => {
		const node = new M365Agent();
		const ctx = makeExecuteContext({
			inputItems: [
				{
					conversationReference: {
						serviceUrl: 'x',
						conversation: { id: 'c' },
						channelId: 'msteams',
					},
				},
			],
			credentials: makeCredentials(),
			parameters: { resource: 'message', operation: 'bogus' },
		});
		await expect(node.execute.call(ctx as unknown as IExecuteFunctions)).rejects.toThrow(
			/Unknown message operation: bogus/,
		);
	});

	it('throws NodeOperationError for an unknown card operation', async () => {
		const node = new M365Agent();
		const ctx = makeExecuteContext({
			inputItems: [
				{
					conversationReference: {
						serviceUrl: 'x',
						conversation: { id: 'c' },
						channelId: 'msteams',
					},
				},
			],
			credentials: makeCredentials(),
			parameters: { resource: 'card', operation: 'bogus' },
		});
		await expect(node.execute.call(ctx as unknown as IExecuteFunctions)).rejects.toThrow(
			/Unknown card operation: bogus/,
		);
	});

	it('throws NodeOperationError for an unknown invokeResponse operation', async () => {
		const node = new M365Agent();
		const ctx = makeExecuteContext({
			inputItems: [{}],
			credentials: makeCredentials(),
			parameters: { resource: 'invokeResponse', operation: 'bogus' },
		});
		await expect(node.execute.call(ctx as unknown as IExecuteFunctions)).rejects.toThrow(
			/Unknown invokeResponse operation: bogus/,
		);
	});
});

// Regression guard for the router's envelope-preservation contract (manifest §12).
// The new mentions / suggestedActions wiring touches only the outbound `activity`,
// never the output spread `{ ...items[i], json: result, pairedItem: i }`. This test
// pins that contract: ancillary input fields must still flow through unchanged when
// the new Options path is exercised.
describe('M365Agent router — envelope preservation with new options', () => {
	it('ancillary item fields survive send-with-mentions', async () => {
		const fakeSend = vi.fn().mockResolvedValue({ id: 'x' });
		(mockCreateConnectorFromBearer as unknown as ReturnType<typeof vi.fn>).mockReturnValue({
			client: {
				sendToConversation: fakeSend,
				replyToActivity: vi.fn(),
				updateActivity: vi.fn(),
				deleteActivity: vi.fn(),
			},
			axios: {},
			token: 't',
			baseURL: 'b',
		});

		const node = new M365Agent();
		const ctx = makeExecuteContext({
			inputItems: [
				{
					conversationReference: {
						serviceUrl: 's',
						conversation: { id: 'c' },
						channelId: 'msteams',
					},
					activity: {},
					customField: 'preserved-through-options-path',
				},
			],
			credentials: makeCredentials(),
			parameters: {
				resource: 'message',
				operation: 'send',
				conversationSource: 'envelope',
				text: 'hi',
				options: {
					mentions: { values: [{ type: 'user', id: '29:u', name: 'U' }] },
				},
			},
		});

		const result = await node.execute.call(ctx as unknown as IExecuteFunctions);
		expect(result[0][0].json).toMatchObject({ customField: 'preserved-through-options-path' });
	});
});
