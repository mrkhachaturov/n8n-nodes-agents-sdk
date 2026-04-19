import { describe, it, expect } from 'vitest';
import { versionDescription } from '../../../nodes/M365Agent/actions/versionDescription';

describe('M365Agent versionDescription', () => {
	it('declares resources in alphabetical order (by display name)', () => {
		// Assert the array as-written — sorting before compare hides misordering.
		const resource = versionDescription.properties.find((p) => p.name === 'resource');
		const names = (resource?.options as { name: string }[] | undefined)?.map((o) => o.name);
		expect(names).toEqual(['Adaptive Card', 'Hero Card', 'Invoke Response', 'Message']);
	});

	it('has a manifest-compliant subtitle', () => {
		expect(versionDescription.subtitle).toBe(
			'={{$parameter["operation"] + ": " + $parameter["resource"]}}',
		);
	});

	it('uses m365AgentApi credential', () => {
		expect(versionDescription.credentials?.[0]?.name).toBe('m365AgentApi');
	});

	it('resource param has noDataExpression', () => {
		const resource = versionDescription.properties.find((p) => p.name === 'resource');
		expect(resource?.noDataExpression).toBe(true);
	});
});
