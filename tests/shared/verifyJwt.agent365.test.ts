import { describe, it, expect } from 'vitest';
import jwt from 'jsonwebtoken';
import { verifyJwt, type VerifyJwtOptions } from '../../shared/verifyJwt';

const secret = 'unit-test-hmac-secret-agent365';

const baseOptions: VerifyJwtOptions = {
	clientId: 'client-id',
	tenantId: 'tid',
	authority: 'https://login.microsoftonline.com',
	getSigningKey: async () => secret,
	verifyAlgorithms: ['HS256'],
};

function makeToken(overrides: Partial<jwt.JwtPayload> = {}): string {
	const payload: jwt.JwtPayload = {
		iss: 'https://api.botframework.com',
		aud: 'client-id',
		exp: Math.floor(Date.now() / 1000) + 60,
		...overrides,
	};
	return jwt.sign(payload, secret, { algorithm: 'HS256' });
}

describe('verifyJwt — audiences/issuers overrides', () => {
	describe('audiences unset (classic defaults)', () => {
		it('accepts token with aud=clientId', async () => {
			const token = makeToken({ aud: 'client-id' });
			const payload = await verifyJwt(token, baseOptions);
			expect(payload.aud).toBe('client-id');
		});

		it('accepts token with aud=https://api.botframework.com', async () => {
			const token = makeToken({ aud: 'https://api.botframework.com' });
			const payload = await verifyJwt(token, baseOptions);
			expect(payload.aud).toBe('https://api.botframework.com');
		});

		it('rejects token with aud=other-audience', async () => {
			const token = makeToken({ aud: 'other-audience' });
			await expect(verifyJwt(token, baseOptions)).rejects.toThrow(/audience/i);
		});
	});

	describe('audiences: ["X"] override', () => {
		const opts: VerifyJwtOptions = { ...baseOptions, audiences: ['X'] };

		it('accepts token with aud=X', async () => {
			const token = makeToken({ aud: 'X' });
			const payload = await verifyJwt(token, opts);
			expect(payload.aud).toBe('X');
		});

		it('rejects token with aud=client-id (replaced by override)', async () => {
			const token = makeToken({ aud: 'client-id' });
			await expect(verifyJwt(token, opts)).rejects.toThrow(/audience/i);
		});

		it('rejects token with aud=https://api.botframework.com (replaced by override)', async () => {
			const token = makeToken({ aud: 'https://api.botframework.com' });
			await expect(verifyJwt(token, opts)).rejects.toThrow(/audience/i);
		});
	});

	describe('issuers override', () => {
		const allowedIssuer = 'https://login.microsoftonline.com/tid/v2.0';
		const opts: VerifyJwtOptions = {
			...baseOptions,
			issuers: [allowedIssuer],
		};

		it('accepts token with matching issuer', async () => {
			const token = makeToken({ iss: allowedIssuer });
			const payload = await verifyJwt(token, opts);
			expect(payload.iss).toBe(allowedIssuer);
		});

		it('rejects token with mismatched issuer', async () => {
			const token = makeToken({ iss: 'https://other-issuer.example.com/v2.0' });
			await expect(verifyJwt(token, opts)).rejects.toThrow(/issuer|jwt/i);
		});
	});
});
