import { describe, it, expect, vi, beforeEach } from 'vitest';
import jwt from 'jsonwebtoken';
import { verifyJwt, type VerifyJwtOptions } from '../../shared/verifyJwt';

const secret = 'unit-test-hmac-secret';

function makeToken(overrides: Partial<jwt.JwtPayload> = {}, opts: jwt.SignOptions = {}): string {
	const payload: jwt.JwtPayload = {
		iss: 'https://api.botframework.com',
		aud: 'client-id',
		exp: Math.floor(Date.now() / 1000) + 60,
		...overrides,
	};
	return jwt.sign(payload, secret, { algorithm: 'HS256', ...opts });
}

const options: VerifyJwtOptions = {
	clientId: 'client-id',
	tenantId: 'tid',
	authority: 'https://login.microsoftonline.com',
	getSigningKey: async () => secret,
	verifyAlgorithms: ['HS256'],
};

describe('shared/verifyJwt', () => {
	it('accepts a valid token', async () => {
		const token = makeToken();
		const payload = await verifyJwt(token, options);
		expect(payload.aud).toBe('client-id');
	});

	it('rejects an expired token', async () => {
		const token = makeToken({ exp: Math.floor(Date.now() / 1000) - 1000 });
		await expect(verifyJwt(token, options)).rejects.toThrow(/jwt expired|expired/i);
	});

	it('rejects a token with the wrong audience', async () => {
		const token = makeToken({ aud: 'other' });
		await expect(verifyJwt(token, options)).rejects.toThrow(/audience/i);
	});

	it('rejects an undecodable token', async () => {
		await expect(verifyJwt('not-a-jwt', options)).rejects.toThrow();
	});

	it('accepts the botframework.com API audience as well as clientId', async () => {
		const token = makeToken({ aud: 'https://api.botframework.com' });
		const payload = await verifyJwt(token, options);
		expect(payload.aud).toBe('https://api.botframework.com');
	});

	it('builds the correct JWKS URI for Bot Framework issuer', async () => {
		const spy = vi.fn(async () => secret);
		const token = makeToken();
		await verifyJwt(token, { ...options, getSigningKey: spy });
		expect(spy).toHaveBeenCalledWith(
			'https://login.botframework.com/v1/.well-known/keys',
			expect.any(Object),
		);
	});

	it('builds the correct JWKS URI for Entra issuer, using tenant from iss', async () => {
		const spy = vi.fn(async () => secret);
		const token = makeToken({ iss: 'https://login.microsoftonline.com/contoso-tenant-id/v2.0' });
		await verifyJwt(token, { ...options, getSigningKey: spy });
		expect(spy).toHaveBeenCalledWith(
			'https://login.microsoftonline.com/contoso-tenant-id/discovery/v2.0/keys',
			expect.any(Object),
		);
	});

	it('uses iss-derived tenant over configured tenantId (multi-tenant scenario)', async () => {
		const spy = vi.fn(async () => secret);
		const token = makeToken({ iss: 'https://login.microsoftonline.com/real-tid/v2.0' });
		await verifyJwt(token, { ...options, tenantId: 'configured-tid', getSigningKey: spy });
		expect(spy).toHaveBeenCalledWith(
			'https://login.microsoftonline.com/real-tid/discovery/v2.0/keys',
			expect.any(Object),
		);
	});

	it('falls back to configured tenantId if iss tenant is "common"', async () => {
		const spy = vi.fn(async () => secret);
		const token = makeToken({ iss: 'https://login.microsoftonline.com/common/v2.0' });
		await verifyJwt(token, { ...options, tenantId: 'fallback-tid', getSigningKey: spy });
		expect(spy).toHaveBeenCalledWith(
			'https://login.microsoftonline.com/fallback-tid/discovery/v2.0/keys',
			expect.any(Object),
		);
	});
});

beforeEach(() => {
	vi.useRealTimers();
});
