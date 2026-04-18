import { describe, it, expect, vi, beforeEach } from 'vitest';
import { acquireOutboundToken } from '../../../shared/auth/router';
import type { M365Agent365Cred } from '../../../shared/types';
import { NodeOperationError } from 'n8n-workflow';

vi.mock('@microsoft/agents-hosting', () => ({
	MsalTokenProvider: vi.fn().mockImplementation(function () {
		return {
			getAccessToken: vi.fn().mockResolvedValue('test-agent365-token'),
		};
	}),
}));

describe('router — agent365 inline backend', () => {
	const baseCred: M365Agent365Cred = {
		tenantId: 't',
		blueprintAppId: 'bp',
		transport: 'inline',
		inlineCredKind: 'clientSecret',
		blueprintSecret: 'sec',
		validateVia: 'sameAsOutbound',
	};

	beforeEach(() => vi.clearAllMocks());

	it('autonomous mode acquires Messaging Bot API token', async () => {
		const result = await acquireOutboundToken({
			authKind: 'agent365',
			credentials: baseCred,
			identityMode: 'autonomous',
			downstreamApi: 'MessagingBotApi',
		});
		expect(result.authorizationHeader).toBe('Bearer test-agent365-token');
	});

	it('agentUser + inline rejects with NodeOperationError', async () => {
		await expect(
			acquireOutboundToken({
				authKind: 'agent365',
				credentials: baseCred,
				identityMode: 'agentUser',
				downstreamApi: 'MessagingBotApi',
				agentUsername: 'x@y.com',
			}),
		).rejects.toThrow(/requires sidecar transport/i);
	});

	it('interactiveOBO + inline rejects with NodeOperationError', async () => {
		await expect(
			acquireOutboundToken({
				authKind: 'agent365',
				credentials: baseCred,
				identityMode: 'interactiveOBO',
				downstreamApi: 'Graph',
				inboundBearer: 'Bearer abc',
			}),
		).rejects.toThrow(/not implemented in M1/i);
	});
});
