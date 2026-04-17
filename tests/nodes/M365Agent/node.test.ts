import { describe, it, expect, vi } from 'vitest';
import { M365Agent } from '../../../nodes/M365Agent/M365Agent.node';
import { makeExecuteContext, makeCredentials } from '../../helpers/makeContext';
import type { IExecuteFunctions } from 'n8n-workflow';

vi.mock('../../../shared/botConnector', () => ({
	createConnector: vi.fn(),
	replyInThread: vi.fn(),
}));

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
			inputItems: [{ conversationReference: { serviceUrl: 'x', conversation: { id: 'c' }, channelId: 'msteams' } }],
			credentials: makeCredentials(),
			parameters: { resource: 'message', operation: 'bogus' },
		});
		await expect(
			node.execute.call(ctx as unknown as IExecuteFunctions),
		).rejects.toThrow(/Unknown message operation: bogus/);
	});

	it('throws NodeOperationError for an unknown card operation', async () => {
		const node = new M365Agent();
		const ctx = makeExecuteContext({
			inputItems: [{ conversationReference: { serviceUrl: 'x', conversation: { id: 'c' }, channelId: 'msteams' } }],
			credentials: makeCredentials(),
			parameters: { resource: 'card', operation: 'bogus' },
		});
		await expect(
			node.execute.call(ctx as unknown as IExecuteFunctions),
		).rejects.toThrow(/Unknown card operation: bogus/);
	});

	it('throws NodeOperationError for an unknown invokeResponse operation', async () => {
		const node = new M365Agent();
		const ctx = makeExecuteContext({
			inputItems: [{}],
			credentials: makeCredentials(),
			parameters: { resource: 'invokeResponse', operation: 'bogus' },
		});
		await expect(
			node.execute.call(ctx as unknown as IExecuteFunctions),
		).rejects.toThrow(/Unknown invokeResponse operation: bogus/);
	});
});
