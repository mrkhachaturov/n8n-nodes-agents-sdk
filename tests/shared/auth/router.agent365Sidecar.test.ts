import { describe, it, expect, vi, beforeEach } from 'vitest';
import { acquireOutboundToken } from '../../../shared/auth/router';
import type { M365Agent365Cred } from '../../../shared/types';

const mockFetch = vi.fn();
globalThis.fetch = mockFetch as any;

describe('router — agent365 sidecar backend', () => {
  const cred: M365Agent365Cred = {
    tenantId: 't',
    blueprintAppId: 'bp-123',
    transport: 'sidecar',
    sidecarUrl: 'http://sidecar:5000',
    outboundDownstreamApi: 'MessagingBotApi',
    validateVia: 'sameAsOutbound',
  };

  beforeEach(() => {
    mockFetch.mockReset();
    mockFetch.mockResolvedValue({
      ok: true,
      status: 200,
      json: () => Promise.resolve({ authorizationHeader: 'Bearer xyz' }),
    });
  });

  it('autonomous calls /AuthorizationHeaderUnauthenticated with AgentIdentity', async () => {
    await acquireOutboundToken({
      authKind: 'agent365',
      credentials: cred,
      identityMode: 'autonomous',
      downstreamApi: 'MessagingBotApi',
    });
    const url = mockFetch.mock.calls[0][0] as string;
    expect(url).toContain('/AuthorizationHeaderUnauthenticated/MessagingBotApi');
    expect(url).toContain('AgentIdentity=bp-123');
  });

  it('agentUser adds AgentUsername param', async () => {
    await acquireOutboundToken({
      authKind: 'agent365',
      credentials: cred,
      identityMode: 'agentUser',
      downstreamApi: 'MessagingBotApi',
      agentUsername: 'agent@contoso.com',
    });
    const url = mockFetch.mock.calls[0][0] as string;
    expect(url).toContain('AgentUsername=agent%40contoso.com');
  });

  it('agentUser with userId adds AgentUserId param', async () => {
    await acquireOutboundToken({
      authKind: 'agent365',
      credentials: cred,
      identityMode: 'agentUser',
      downstreamApi: 'MessagingBotApi',
      agentUserId: 'oid-guid-here',
    });
    const url = mockFetch.mock.calls[0][0] as string;
    expect(url).toContain('AgentUserId=oid-guid-here');
    expect(url).not.toContain('AgentUsername');
  });

  it('interactiveOBO uses /AuthorizationHeader with forwarded bearer', async () => {
    await acquireOutboundToken({
      authKind: 'agent365',
      credentials: cred,
      identityMode: 'interactiveOBO',
      downstreamApi: 'Graph',
      inboundBearer: 'Bearer user-token',
    });
    const url = mockFetch.mock.calls[0][0] as string;
    const opts = mockFetch.mock.calls[0][1] as RequestInit;
    expect(url).toContain('/AuthorizationHeader/Graph');
    expect((opts.headers as Record<string, string>).Authorization).toBe('Bearer user-token');
  });

  it('returns the authorizationHeader from the response', async () => {
    const result = await acquireOutboundToken({
      authKind: 'agent365',
      credentials: cred,
      identityMode: 'autonomous',
      downstreamApi: 'MessagingBotApi',
    });
    expect(result.authorizationHeader).toBe('Bearer xyz');
  });

  it('sidecar non-2xx response throws with cause', async () => {
    mockFetch.mockResolvedValueOnce({
      ok: false,
      status: 500,
      text: () => Promise.resolve('sidecar error'),
    });
    await expect(acquireOutboundToken({
      authKind: 'agent365',
      credentials: cred,
      identityMode: 'autonomous',
      downstreamApi: 'MessagingBotApi',
    })).rejects.toThrow(/sidecar.*500/i);
  });
});
