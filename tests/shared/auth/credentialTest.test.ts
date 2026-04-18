import { describe, it, expect, vi, beforeEach } from 'vitest';
import { agent365CredentialTest } from '../../../shared/auth/credentialTest';

const mockFetch = vi.fn();
globalThis.fetch = mockFetch as any;

// Hoist the mock handle so vi.mock() factory can use it.
const { mockMsalGetAccessToken } = vi.hoisted(() => ({
	mockMsalGetAccessToken: vi.fn(),
}));

vi.mock('@microsoft/agents-hosting', () => ({
	MsalTokenProvider: vi.fn().mockImplementation(function () {
		return { getAccessToken: mockMsalGetAccessToken };
	}),
}));

describe('agent365CredentialTest', () => {
	beforeEach(() => {
		mockFetch.mockReset();
		mockMsalGetAccessToken.mockReset();
	});

	it('inline + clientSecret → OK when MSAL returns a token', async () => {
		mockMsalGetAccessToken.mockResolvedValueOnce('tok');
		const r = await agent365CredentialTest.call({} as any, {
			data: {
				tenantId: 't', blueprintAppId: 'bp', transport: 'inline',
				inlineCredKind: 'clientSecret', blueprintSecret: 's',
				validateVia: 'sameAsOutbound', outboundDownstreamApi: 'MessagingBotApi',
			},
		} as any);
		expect(r).toEqual({ status: 'OK', message: expect.stringContaining('inline') });
	});

	it('sidecar → OK when /healthz and /AuthorizationHeaderUnauthenticated both 200', async () => {
		mockFetch
			.mockResolvedValueOnce({ ok: true, status: 200 })
			.mockResolvedValueOnce({ ok: true, json: () => Promise.resolve({ authorizationHeader: 'Bearer x' }) });
		const r = await agent365CredentialTest.call({} as any, {
			data: {
				tenantId: 't', blueprintAppId: 'bp', transport: 'sidecar',
				sidecarUrl: 'http://s:5000', outboundDownstreamApi: 'MessagingBotApi',
				validateVia: 'sameAsOutbound',
			},
		} as any);
		expect(r.status).toBe('OK');
	});

	it('sidecar /healthz failure → Error', async () => {
		mockFetch.mockResolvedValueOnce({ ok: false, status: 503 });
		const r = await agent365CredentialTest.call({} as any, {
			data: {
				tenantId: 't', blueprintAppId: 'bp', transport: 'sidecar',
				sidecarUrl: 'http://s:5000', outboundDownstreamApi: 'MessagingBotApi',
				validateVia: 'sameAsOutbound',
			},
		} as any);
		expect(r.status).toBe('Error');
		expect(r.message).toMatch(/healthz|503/i);
	});
});
