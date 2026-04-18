import { describe, it, expect } from 'vitest';
import { description as ir } from '../../../nodes/M365Agent/actions/invokeResponse/respond.operation';

describe('Invoke Response: respond — unchanged in v0.3.0', () => {
	it('has no identityMode field', () => {
		expect(ir.find((p: any) => p.name === 'identityMode')).toBeUndefined();
	});

	it('has no agentUsername field', () => {
		expect(ir.find((p: any) => p.name === 'agentUsername')).toBeUndefined();
	});
});
