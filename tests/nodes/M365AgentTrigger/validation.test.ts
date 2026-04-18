import { describe, it, expect, vi, beforeEach } from 'vitest';
import { M365AgentTrigger } from '../../../nodes/M365AgentTrigger/M365AgentTrigger.node';
import { makeWebhookMock } from '../../helpers/webhookMock';

vi.mock('../../../shared/auth/router', () => ({
  validateInboundToken: vi.fn(),
}));
import { validateInboundToken } from '../../../shared/auth/router';

const SAMPLE_ACTIVITY = {
  type: 'message', id: 'a1', serviceUrl: 'https://smba/', channelId: 'msteams',
  from: { id: '29:u1', name: 'U' },
  conversation: { id: '19:c1', conversationType: 'personal' },
  recipient: { id: '28:b1' }, text: 'hi',
};

describe('M365AgentTrigger webhook — validation routing', () => {
  beforeEach(() => (validateInboundToken as any).mockReset());

  it('classic → router called with classicBot authKind, no authContext in envelope', async () => {
    (validateInboundToken as any).mockResolvedValue({ claims: { aud: 'app-id' } });
    const ctx = makeWebhookMock({
      body: SAMPLE_ACTIVITY,
      headers: { authorization: 'Bearer eyJ-classic-token' },
      nodeParams: { authKind: 'classicBot', responseMode: 'immediate' },
      credentialName: 'm365AgentApi',
      credentials: { appType: 'SingleTenant', clientId: 'app-id', tenantId: 't', clientSecret: 's' },
    });
    const trigger = new M365AgentTrigger();
    const response = await trigger.webhook!.call(ctx);
    const workflowData = (response as any).workflowData as Array<Array<{ json: any }>>;
    const envelope = workflowData[0][0].json;
    expect(envelope.authContext).toBeUndefined();
    expect(validateInboundToken).toHaveBeenCalledWith('classicBot', expect.anything(), 'Bearer eyJ-classic-token');
  });

  it('agent365 → router called with agent365, authContext populated', async () => {
    (validateInboundToken as any).mockResolvedValue({
      claims: { aud: 'a6c6ce43-d3ed-40ae-a6d2-4f9c028e0355', iss: 'https://login.microsoftonline.com/t/', oid: 'user-oid' },
    });
    const ctx = makeWebhookMock({
      body: SAMPLE_ACTIVITY,
      headers: { authorization: 'Bearer eyJ-agent365-token' },
      nodeParams: { authKind: 'agent365', responseMode: 'immediate' },
      credentialName: 'm365Agent365Api',
      credentials: {
        tenantId: 't', blueprintAppId: 'bp', transport: 'sidecar',
        sidecarUrl: 'http://s:5000', validateVia: 'sameAsOutbound',
      },
    });
    const trigger = new M365AgentTrigger();
    const response = await trigger.webhook!.call(ctx);
    const workflowData = (response as any).workflowData as Array<Array<{ json: any }>>;
    const envelope = workflowData[0][0].json;
    expect(envelope.authContext).toBeDefined();
    expect(envelope.authContext.inboundBearer).toBe('Bearer eyJ-agent365-token');
    expect(envelope.authContext.validatedClaims.oid).toBe('user-oid');
    expect(envelope.authContext.tokenSource).toBe('messagingBotApi');
    expect(validateInboundToken).toHaveBeenCalledWith('agent365', expect.anything(), 'Bearer eyJ-agent365-token');
  });
});
