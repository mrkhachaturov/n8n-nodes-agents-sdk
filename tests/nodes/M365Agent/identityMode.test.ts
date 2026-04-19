import { describe, it, expect } from 'vitest';
import { description as sendDesc } from '../../../nodes/M365Agent/actions/message/send.operation';
import { description as cardSendDesc } from '../../../nodes/M365Agent/actions/card-adaptive/send.operation';
import { description as cardUpdateDesc } from '../../../nodes/M365Agent/actions/card-adaptive/update.operation';

describe('Message operations — identityMode field', () => {
	it('send has identityMode gated on authKind=agent365', () => {
		const idm = sendDesc.find((p: any) => p.name === 'identityMode')!;
		expect(idm).toBeDefined();
		expect(idm.displayOptions.show.authKind).toEqual(['agent365']);
		expect(idm.displayOptions.show.resource).toEqual(['message']);
		expect(idm.displayOptions.show.operation).toEqual(['send']);
		expect(idm.noDataExpression).toBe(true);
		expect(idm.default).toBe('autonomous');
		expect(idm.options.map((o: any) => o.value).sort()).toEqual(['agentUser', 'autonomous']);
	});

	it('send has agentUsername field gated on identityMode=agentUser', () => {
		const u = sendDesc.find((p: any) => p.name === 'agentUsername')!;
		expect(u.displayOptions.show.identityMode).toEqual(['agentUser']);
	});
});

describe('Card operations — identityMode field', () => {
	it('card send has identityMode gated on authKind=agent365 + resource=adaptiveCard + operation=send', () => {
		const idm = cardSendDesc.find((p: any) => p.name === 'identityMode')!;
		expect(idm).toBeDefined();
		expect(idm.displayOptions.show.authKind).toEqual(['agent365']);
		expect(idm.displayOptions.show.resource).toEqual(['adaptiveCard']);
		expect(idm.displayOptions.show.operation).toEqual(['send']);
		expect(idm.default).toBe('autonomous');
	});

	it('card send has agentUsername field gated on identityMode=agentUser', () => {
		const u = cardSendDesc.find((p: any) => p.name === 'agentUsername')!;
		expect(u).toBeDefined();
		expect(u.displayOptions.show.identityMode).toEqual(['agentUser']);
	});

	it('card update has identityMode gated on authKind=agent365 + resource=adaptiveCard + operation=update', () => {
		const idm = cardUpdateDesc.find((p: any) => p.name === 'identityMode')!;
		expect(idm).toBeDefined();
		expect(idm.displayOptions.show.authKind).toEqual(['agent365']);
		expect(idm.displayOptions.show.resource).toEqual(['adaptiveCard']);
		expect(idm.displayOptions.show.operation).toEqual(['update']);
		expect(idm.default).toBe('autonomous');
	});

	it('card update has agentUsername field gated on identityMode=agentUser', () => {
		const u = cardUpdateDesc.find((p: any) => p.name === 'agentUsername')!;
		expect(u).toBeDefined();
		expect(u.displayOptions.show.identityMode).toEqual(['agentUser']);
	});
});
