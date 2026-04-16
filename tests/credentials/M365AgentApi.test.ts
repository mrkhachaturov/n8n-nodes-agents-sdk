import { describe, it, expect } from 'vitest';
import { M365AgentApi } from '../../credentials/M365AgentApi.credentials';

describe('M365AgentApi credential', () => {
	it('registers the expected name', () => {
		const c = new M365AgentApi();
		expect(c.name).toBe('m365AgentApi');
	});

	it('exposes all required fields', () => {
		const c = new M365AgentApi();
		const names = c.properties.map((p) => p.name).sort();
		expect(names).toEqual(
			['anonymousAllowed', 'appType', 'clientId', 'clientSecret', 'tenantId'].sort(),
		);
	});

	it('clientSecret is flagged as password', () => {
		const c = new M365AgentApi();
		const cs = c.properties.find((p) => p.name === 'clientSecret');
		expect(cs?.typeOptions?.password).toBe(true);
	});

	it('appType has three options', () => {
		const c = new M365AgentApi();
		const appType = c.properties.find((p) => p.name === 'appType');
		expect(appType?.options).toHaveLength(3);
	});
});
