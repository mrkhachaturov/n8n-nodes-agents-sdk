import { describe, it, expect } from 'vitest';
import { description as sendDesc } from '../../../nodes/M365Agent/actions/message/send.operation';

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
