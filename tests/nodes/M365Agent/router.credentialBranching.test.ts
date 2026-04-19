import { describe, it, expect, vi, beforeEach } from 'vitest';

vi.mock('../../../nodes/M365Agent/actions/message', () => ({
	send: { execute: vi.fn().mockResolvedValue({ id: 'stub-send' }) },
	reply: { execute: vi.fn().mockResolvedValue({ id: 'stub-reply' }) },
	update: { execute: vi.fn().mockResolvedValue({ id: 'stub-update' }) },
	deleteMessage: { execute: vi.fn().mockResolvedValue({ id: 'stub-delete' }) },
	replyInThread: { execute: vi.fn().mockResolvedValue({ id: 'stub-thread' }) },
}));
vi.mock('../../../nodes/M365Agent/actions/card-adaptive', () => ({
	send: { execute: vi.fn().mockResolvedValue({ id: 'stub-card-send' }) },
	update: { execute: vi.fn().mockResolvedValue({ id: 'stub-card-update' }) },
}));
vi.mock('../../../nodes/M365Agent/actions/invokeResponse', () => ({
	respond: { execute: vi.fn().mockResolvedValue(undefined) },
}));

import { router } from '../../../nodes/M365Agent/actions/router';

describe('actions/router — credential branching + bundle cache key', () => {
	const getCredentials = vi.fn();
	const getNodeParameter = vi.fn();
	const getInputData = vi.fn().mockReturnValue([{ json: {} }]);
	const sendResponse = vi.fn();
	const fakeCtx: any = {
		getCredentials,
		getNodeParameter,
		getInputData,
		sendResponse,
		getNode: () => ({}),
	};

	beforeEach(() => {
		getCredentials.mockReset();
		getNodeParameter.mockReset();
	});

	it('fetches m365AgentApi when authKind=classicBot', async () => {
		getNodeParameter.mockImplementation((name: string) => {
			if (name === 'resource') return 'message';
			if (name === 'operation') return 'send';
			if (name === 'authKind') return 'classicBot';
			return undefined;
		});
		getCredentials.mockResolvedValueOnce({ appType: 'SingleTenant', clientId: 'bot' });
		await router.call(fakeCtx);
		expect(getCredentials).toHaveBeenCalledWith('m365AgentApi');
		expect(getCredentials).not.toHaveBeenCalledWith('m365Agent365Api');
	});

	it('fetches m365Agent365Api when authKind=agent365', async () => {
		getNodeParameter.mockImplementation((name: string) => {
			if (name === 'resource') return 'message';
			if (name === 'operation') return 'send';
			if (name === 'authKind') return 'agent365';
			return undefined;
		});
		getCredentials.mockResolvedValueOnce({
			tenantId: 't',
			blueprintAppId: 'bp',
			transport: 'sidecar',
		});
		await router.call(fakeCtx);
		expect(getCredentials).toHaveBeenCalledWith('m365Agent365Api');
		expect(getCredentials).not.toHaveBeenCalledWith('m365AgentApi');
	});

	it('makeBundleKey separates items with same serviceUrl but different identityMode', async () => {
		const { makeBundleKey } = await import('../../../nodes/M365Agent/actions/router');
		const a = makeBundleKey('https://x/', 'classicBot', 'autonomous');
		const b = makeBundleKey('https://x/', 'agent365', 'autonomous');
		const c = makeBundleKey('https://x/', 'agent365', 'agentUser', 'agent@x.com');
		const d = makeBundleKey('https://x/', 'agent365', 'agentUser', 'other@x.com');
		const keys = new Set([a, b, c, d]);
		expect(keys.size).toBe(4);
	});
});
