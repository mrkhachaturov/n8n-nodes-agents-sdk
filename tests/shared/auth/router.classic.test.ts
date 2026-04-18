import { describe, it, expect, vi, beforeEach } from 'vitest';
import { acquireOutboundToken, validateInboundToken } from '../../../shared/auth/router';
import type { M365ClassicBotCred } from '../../../shared/types';

vi.mock('@microsoft/agents-hosting', () => ({
	MsalTokenProvider: vi.fn().mockImplementation(function () {
		return { getAccessToken: vi.fn().mockResolvedValue('test-bot-token') };
	}),
}));

describe('router — classic bot path', () => {
	const cred: M365ClassicBotCred = {
		appType: 'SingleTenant',
		clientId: 'app-id',
		clientSecret: 'secret',
		tenantId: 'tenant',
	};

	beforeEach(() => {
		vi.clearAllMocks();
	});

	it('acquires bot framework token for classic authKind', async () => {
		const result = await acquireOutboundToken({
			authKind: 'classicBot',
			credentials: cred,
			identityMode: 'autonomous',
			downstreamApi: 'BotFramework',
		});
		expect(result.authorizationHeader).toBe('Bearer test-bot-token');
	});

	it('ignores identityMode when authKind is classic', async () => {
		const result = await acquireOutboundToken({
			authKind: 'classicBot',
			credentials: cred,
			identityMode: 'agentUser' as any,
			downstreamApi: 'BotFramework',
		});
		expect(result.authorizationHeader).toBe('Bearer test-bot-token');
	});
});
