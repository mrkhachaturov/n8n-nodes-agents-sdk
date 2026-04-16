import { describe, it, expect } from 'vitest';
import jwt from 'jsonwebtoken';
import { buildAuthConfig } from '../../shared/buildAuthConfig';
import { verifyJwt } from '../../shared/verifyJwt';
import type { M365AgentCredentials } from '../../shared/types';
import { MsalTokenProvider } from '@microsoft/agents-hosting';

const clientId = process.env.M365_TEST_CLIENT_ID;
const clientSecret = process.env.M365_TEST_CLIENT_SECRET;
const tenantId = process.env.M365_TEST_TENANT_ID;

const hasRealCreds = Boolean(clientId && clientSecret && tenantId);
const descr = hasRealCreds ? describe : describe.skip;

descr('Azure integration (real credentials)', () => {
	const creds: M365AgentCredentials = {
		appType: 'SingleTenant',
		clientId: clientId as string,
		clientSecret: clientSecret as string,
		tenantId: tenantId as string,
	};

	it('acquires a real Bot Framework access token via MSAL', async () => {
		const authConfig = buildAuthConfig(creds);
		const provider = new MsalTokenProvider();
		const token = await provider.getAccessToken(authConfig, 'https://api.botframework.com');

		expect(typeof token).toBe('string');
		expect(token.length).toBeGreaterThan(100);

		const decoded = jwt.decode(token) as jwt.JwtPayload | null;
		expect(decoded).not.toBeNull();
		expect(decoded?.aud).toBe('https://api.botframework.com');
		expect(decoded?.appid ?? decoded?.azp).toBe(clientId);
	}, 30_000);

	it('the issued token passes verifyJwt against our JWKS logic', async () => {
		const authConfig = buildAuthConfig(creds);
		const provider = new MsalTokenProvider();
		const token = await provider.getAccessToken(authConfig, 'https://api.botframework.com');

		// Tokens Azure issues to the bot itself have `aud: https://api.botframework.com`
		// and iss containing the tenant — the same verifier used for inbound
		// activities should accept this token shape.
		const payload = await verifyJwt(token, {
			clientId: clientId as string,
			tenantId: tenantId as string,
		});

		expect(payload.aud).toBe('https://api.botframework.com');
		expect(payload.iss).toMatch(new RegExp(tenantId as string));
	}, 30_000);
});
