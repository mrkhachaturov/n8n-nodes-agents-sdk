import { describe, it, expect } from 'vitest';
import { buildAuthConfig } from '../../shared/buildAuthConfig';
import type { M365AgentCredentials } from '../../shared/types';

describe('shared/buildAuthConfig', () => {
	it('maps single-tenant credentials to an AuthConfiguration', () => {
		const creds: M365AgentCredentials = {
			appType: 'SingleTenant',
			clientId: 'cid',
			clientSecret: 'sec',
			tenantId: 'tid',
		};
		const cfg = buildAuthConfig(creds);
		expect(cfg.clientId).toBe('cid');
		expect(cfg.clientSecret).toBe('sec');
		expect(cfg.tenantId).toBe('tid');
		expect(cfg.authority).toBe('https://login.microsoftonline.com');
	});

	it('maps multi-tenant credentials and omits tenantId so "botframework.com" is used', () => {
		const creds: M365AgentCredentials = {
			appType: 'MultiTenant',
			clientId: 'cid',
			clientSecret: 'sec',
		};
		const cfg = buildAuthConfig(creds);
		expect(cfg.clientId).toBe('cid');
		expect(cfg.clientSecret).toBe('sec');
		expect(cfg.tenantId).toBeUndefined();
	});

	it('maps user-assigned managed identity without a clientSecret', () => {
		const creds: M365AgentCredentials = {
			appType: 'UserAssignedMsi',
			clientId: 'cid',
			tenantId: 'tid',
		};
		const cfg = buildAuthConfig(creds);
		expect(cfg.clientId).toBe('cid');
		expect(cfg.tenantId).toBe('tid');
		expect(cfg.clientSecret).toBeUndefined();
	});

	it('throws when single-tenant credentials are missing a tenantId', () => {
		const creds: M365AgentCredentials = {
			appType: 'SingleTenant',
			clientId: 'cid',
			clientSecret: 'sec',
		};
		expect(() => buildAuthConfig(creds)).toThrow(/tenantId/i);
	});

	it('throws when credentials other than UAMI are missing a clientSecret', () => {
		const creds: M365AgentCredentials = {
			appType: 'SingleTenant',
			clientId: 'cid',
			tenantId: 'tid',
		};
		expect(() => buildAuthConfig(creds)).toThrow(/clientSecret/);
	});
});
