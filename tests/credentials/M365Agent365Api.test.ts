import { describe, it, expect } from 'vitest';
import { M365Agent365Api } from '../../credentials/M365Agent365Api.credentials';

describe('M365Agent365Api credential', () => {
	const c = new M365Agent365Api();

	it('has required metadata', () => {
		expect(c.name).toBe('m365Agent365Api');
		expect(c.displayName).toBe('M365 Agent 365 API');
		expect(c.testedBy).toBe('agent365CredentialTest');
	});

	it('declares all Agent 365 fields', () => {
		const names = c.properties.map((p) => p.name);
		expect(names).toContain('tenantId');
		expect(names).toContain('blueprintAppId');
		expect(names).toContain('agentInstanceAppId');
		expect(names).toContain('transport');
		expect(names).toContain('inlineCredKind');
		expect(names).toContain('blueprintSecret');
		expect(names).toContain('blueprintCertPem');
		expect(names).toContain('sidecarUrl');
		expect(names).toContain('outboundDownstreamApi');
		expect(names).toContain('defaultAgentUsername');
		expect(names).toContain('validateVia');
		expect(names).toContain('validatorSidecarUrl');
		expect(names).toContain('options');
	});

	it('spec §6.2 required fields carry `required: true`', () => {
		const mustBeRequired = [
			'tenantId',
			'blueprintAppId',
			'agentInstanceAppId',
			'transport',
			'inlineCredKind',
			'blueprintSecret',
			'blueprintCertPem',
			'blueprintCertThumbprint',
			'sidecarUrl',
			'outboundDownstreamApi',
			'validateVia',
			'validatorSidecarUrl',
		];
		for (const name of mustBeRequired) {
			const prop = c.properties.find((p) => p.name === name);
			expect(prop, `missing property ${name}`).toBeDefined();
			expect(prop!.required, `${name} must be required per spec §6.2`).toBe(true);
		}
	});

	it('optional fields are NOT required', () => {
		for (const name of ['defaultAgentUsername', 'options']) {
			const prop = c.properties.find((p) => p.name === name);
			expect(prop?.required ?? false).toBe(false);
		}
	});

	it('transport field has inline and sidecar options', () => {
		const t = c.properties.find((p) => p.name === 'transport')!;
		expect(t.type).toBe('options');
		expect((t as any).options.map((o: any) => o.value).sort()).toEqual(['inline', 'sidecar']);
	});

	it('blueprintSecret gated on inline + clientSecret', () => {
		const s = c.properties.find((p) => p.name === 'blueprintSecret')!;
		expect(s.displayOptions?.show).toEqual({
			transport: ['inline'],
			inlineCredKind: ['clientSecret'],
		});
	});

	it('agentInstanceAppId gated on sidecar transport', () => {
		const a = c.properties.find((p) => p.name === 'agentInstanceAppId')!;
		expect(a.displayOptions?.show).toEqual({ transport: ['sidecar'] });
	});
});
