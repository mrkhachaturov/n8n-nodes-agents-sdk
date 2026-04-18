import { describe, it, expect, vi, beforeEach } from 'vitest';
import { validateInboundToken } from '../../../shared/auth/router';
import type { M365Agent365Cred, M365ClassicBotCred } from '../../../shared/types';

const { mockVerifyJwt } = vi.hoisted(() => ({ mockVerifyJwt: vi.fn() }));
vi.mock('../../../shared/verifyJwt', () => ({ verifyJwt: mockVerifyJwt }));

const mockFetch = vi.fn();
globalThis.fetch = mockFetch as any;

describe('router — validateInboundToken', () => {
	beforeEach(() => {
		mockVerifyJwt.mockReset();
		mockFetch.mockReset();
	});

	it('classicBot uses existing verifyJwt', async () => {
		mockVerifyJwt.mockResolvedValueOnce({ aud: 'app-id', iss: 'bot-framework-issuer' });
		const cred: M365ClassicBotCred = { appType: 'SingleTenant', clientId: 'app-id' };
		const r = await validateInboundToken('classicBot', cred, 'eyJ...');
		expect(r.claims.aud).toBe('app-id');
		expect(mockVerifyJwt).toHaveBeenCalledTimes(1);
	});

	it('agent365 validateVia=inline extends issuer list', async () => {
		mockVerifyJwt.mockResolvedValueOnce({ aud: 'msgbot-audience' });
		const cred: M365Agent365Cred = {
			tenantId: 't',
			blueprintAppId: 'bp',
			transport: 'sidecar',
			sidecarUrl: 'http://s:5000',
			validateVia: 'inline',
		};
		const r = await validateInboundToken('agent365', cred, 'eyJ...');
		expect(r.claims.aud).toBe('msgbot-audience');
		const call = mockVerifyJwt.mock.calls[0];
		expect(call[1].audiences).toContain('bp');
	});

	it('agent365 default audience list covers both v1 (api://) and v2 (GUID) aud forms', async () => {
		mockVerifyJwt.mockResolvedValue({ aud: 'api://bp' });
		const cred: M365Agent365Cred = {
			tenantId: 't',
			blueprintAppId: 'bp',
			transport: 'sidecar',
			sidecarUrl: 'http://s:5000',
			validateVia: 'inline',
		};
		await validateInboundToken('agent365', cred, 'eyJ...');
		const call = mockVerifyJwt.mock.calls[mockVerifyJwt.mock.calls.length - 1];
		expect(call[1].audiences).toEqual(expect.arrayContaining(['bp', 'api://bp']));
	});

	it('agent365 validateVia=sidecar calls /Validate', async () => {
		mockFetch.mockResolvedValueOnce({
			ok: true,
			json: () => Promise.resolve({ claims: { upn: 'user@x.com' } }),
		});
		const cred: M365Agent365Cred = {
			tenantId: 't',
			blueprintAppId: 'bp',
			transport: 'sidecar',
			sidecarUrl: 'http://s:5000',
			validateVia: 'sidecar',
		};
		const r = await validateInboundToken('agent365', cred, 'Bearer eyJ...');
		expect(r.claims.upn).toBe('user@x.com');
		expect(mockFetch.mock.calls[0][0]).toContain('/Validate');
	});

	it('agent365 validateVia=sameAsOutbound uses transport to decide', async () => {
		mockFetch.mockResolvedValueOnce({
			ok: true,
			json: () => Promise.resolve({ claims: {} }),
		});
		const cred: M365Agent365Cred = {
			tenantId: 't',
			blueprintAppId: 'bp',
			transport: 'sidecar',
			sidecarUrl: 'http://s:5000',
			validateVia: 'sameAsOutbound',
		};
		await validateInboundToken('agent365', cred, 'Bearer eyJ...');
		expect(mockFetch).toHaveBeenCalled();
	});
});
