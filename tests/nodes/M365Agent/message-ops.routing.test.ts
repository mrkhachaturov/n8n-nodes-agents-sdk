import { describe, it, expect, vi, beforeEach } from 'vitest';
import type { BotConnectorBundle } from '../../../shared/botConnector';

vi.mock('../../../shared/auth/router', () => ({
  acquireOutboundToken: vi.fn().mockResolvedValue({ authorizationHeader: 'Bearer routed' }),
}));
vi.mock('../../../shared/botConnector', () => ({
  createConnectorFromBearer: vi.fn().mockImplementation((url: string, header: string) => ({
    client: {
      replyToActivity: vi.fn().mockResolvedValue({ id: 'act-reply' }),
      updateActivity: vi.fn().mockResolvedValue({ id: 'act-update' }),
      deleteActivity: vi.fn().mockResolvedValue(undefined),
    },
    axios: { post: vi.fn().mockResolvedValue({ data: { id: 'act-thread' } }) },
    token: header,
    baseURL: url,
  })),
  replyInThread: vi.fn().mockResolvedValue({ id: 'act-thread' }),
}));

import { execute as replyExecute } from '../../../nodes/M365Agent/actions/message/reply.operation';
import { execute as updateExecute } from '../../../nodes/M365Agent/actions/message/update.operation';
import { execute as deleteExecute } from '../../../nodes/M365Agent/actions/message/deleteMessage.operation';
import { execute as replyInThreadExecute } from '../../../nodes/M365Agent/actions/message/replyInThread.operation';
import { acquireOutboundToken } from '../../../shared/auth/router';
import { makeExecuteMock } from '../../helpers/executeMock';

const ENVELOPE_INPUT = [{
  json: {
    conversationReference: {
      serviceUrl: 'https://smba/',
      conversation: { id: '19:c1', conversationType: 'personal' },
      channelId: 'msteams',
      activityId: 'act-1',
    },
    activity: { type: 'message', text: 'hi' },
  },
}];

// ── reply ─────────────────────────────────────────────────────────────────────

describe('message/reply routing', () => {
  beforeEach(() => (acquireOutboundToken as any).mockClear());

  it('classic authKind → router receives classicBot', async () => {
    const ctx = makeExecuteMock({
      items: ENVELOPE_INPUT,
      nodeParams: (name: string, _i: number, fallback?: unknown) => {
        const map: Record<string, unknown> = {
          authKind: 'classicBot',
          resource: 'message', operation: 'reply',
          conversationSource: 'envelope',
          text: 'hi',
          options: {},
        };
        return map[name] ?? fallback;
      },
      credentialName: 'm365AgentApi',
      credentials: { appType: 'SingleTenant', clientId: 'c', tenantId: 't', clientSecret: 's' },
    });
    const bundles = new Map<string, BotConnectorBundle>();
    await replyExecute.call(ctx, 0, 'classicBot' as any, { appType: 'SingleTenant', clientId: 'c' } as any, bundles);
    expect(acquireOutboundToken).toHaveBeenCalledWith(expect.objectContaining({ authKind: 'classicBot' }));
  });

  it('agent365 + autonomous → router receives downstreamApi=MessagingBotApi', async () => {
    const ctx = makeExecuteMock({
      items: ENVELOPE_INPUT,
      nodeParams: (name: string, _i: number, fallback?: unknown) => {
        const map: Record<string, unknown> = {
          authKind: 'agent365', identityMode: 'autonomous',
          resource: 'message', operation: 'reply',
          conversationSource: 'envelope', text: 'hi', options: {},
        };
        return map[name] ?? fallback;
      },
      credentialName: 'm365Agent365Api',
      credentials: {
        tenantId: 't', blueprintAppId: 'bp', transport: 'sidecar',
        sidecarUrl: 'http://s:5000', outboundDownstreamApi: 'MessagingBotApi',
        validateVia: 'sameAsOutbound',
      },
    });
    await replyExecute.call(ctx, 0, 'agent365' as any, {
      tenantId: 't', blueprintAppId: 'bp', transport: 'sidecar',
      sidecarUrl: 'http://s:5000', outboundDownstreamApi: 'MessagingBotApi',
      validateVia: 'sameAsOutbound',
    } as any, new Map());
    expect(acquireOutboundToken).toHaveBeenCalledWith(expect.objectContaining({
      authKind: 'agent365', identityMode: 'autonomous', downstreamApi: 'MessagingBotApi',
    }));
  });

  it('agent365 + agentUser by UPN → forwards agentUsername', async () => {
    const ctx = makeExecuteMock({
      items: ENVELOPE_INPUT,
      nodeParams: (name: string, _i: number, fallback?: unknown) => {
        const map: Record<string, unknown> = {
          authKind: 'agent365', identityMode: 'agentUser',
          userSelectorMode: 'byUpn', agentUsername: 'a@b.com',
          resource: 'message', operation: 'reply',
          conversationSource: 'envelope', text: 'hi', options: {},
        };
        return map[name] ?? fallback;
      },
      credentialName: 'm365Agent365Api',
      credentials: {
        tenantId: 't', blueprintAppId: 'bp', transport: 'sidecar',
        sidecarUrl: 'http://s:5000', outboundDownstreamApi: 'MessagingBotApi',
        validateVia: 'sameAsOutbound',
      },
    });
    await replyExecute.call(ctx, 0, 'agent365' as any, {
      tenantId: 't', blueprintAppId: 'bp', transport: 'sidecar',
      sidecarUrl: 'http://s:5000', outboundDownstreamApi: 'MessagingBotApi',
      validateVia: 'sameAsOutbound',
    } as any, new Map());
    expect(acquireOutboundToken).toHaveBeenCalledWith(expect.objectContaining({
      identityMode: 'agentUser', agentUsername: 'a@b.com',
    }));
  });
});

// ── update ────────────────────────────────────────────────────────────────────

describe('message/update routing', () => {
  beforeEach(() => (acquireOutboundToken as any).mockClear());

  it('classic authKind → router receives classicBot', async () => {
    const ctx = makeExecuteMock({
      items: ENVELOPE_INPUT,
      nodeParams: (name: string, _i: number, fallback?: unknown) => {
        const map: Record<string, unknown> = {
          authKind: 'classicBot',
          resource: 'message', operation: 'update',
          conversationSource: 'envelope',
          text: 'updated', options: {},
        };
        return map[name] ?? fallback;
      },
      credentialName: 'm365AgentApi',
      credentials: { appType: 'SingleTenant', clientId: 'c', tenantId: 't', clientSecret: 's' },
    });
    const bundles = new Map<string, BotConnectorBundle>();
    await updateExecute.call(ctx, 0, 'classicBot' as any, { appType: 'SingleTenant', clientId: 'c' } as any, bundles);
    expect(acquireOutboundToken).toHaveBeenCalledWith(expect.objectContaining({ authKind: 'classicBot' }));
  });

  it('agent365 + autonomous → router receives downstreamApi=MessagingBotApi', async () => {
    const ctx = makeExecuteMock({
      items: ENVELOPE_INPUT,
      nodeParams: (name: string, _i: number, fallback?: unknown) => {
        const map: Record<string, unknown> = {
          authKind: 'agent365', identityMode: 'autonomous',
          resource: 'message', operation: 'update',
          conversationSource: 'envelope', text: 'updated', options: {},
        };
        return map[name] ?? fallback;
      },
      credentialName: 'm365Agent365Api',
      credentials: {
        tenantId: 't', blueprintAppId: 'bp', transport: 'sidecar',
        sidecarUrl: 'http://s:5000', outboundDownstreamApi: 'MessagingBotApi',
        validateVia: 'sameAsOutbound',
      },
    });
    await updateExecute.call(ctx, 0, 'agent365' as any, {
      tenantId: 't', blueprintAppId: 'bp', transport: 'sidecar',
      sidecarUrl: 'http://s:5000', outboundDownstreamApi: 'MessagingBotApi',
      validateVia: 'sameAsOutbound',
    } as any, new Map());
    expect(acquireOutboundToken).toHaveBeenCalledWith(expect.objectContaining({
      authKind: 'agent365', identityMode: 'autonomous', downstreamApi: 'MessagingBotApi',
    }));
  });

  it('agent365 + agentUser by UPN → forwards agentUsername', async () => {
    const ctx = makeExecuteMock({
      items: ENVELOPE_INPUT,
      nodeParams: (name: string, _i: number, fallback?: unknown) => {
        const map: Record<string, unknown> = {
          authKind: 'agent365', identityMode: 'agentUser',
          userSelectorMode: 'byUpn', agentUsername: 'a@b.com',
          resource: 'message', operation: 'update',
          conversationSource: 'envelope', text: 'updated', options: {},
        };
        return map[name] ?? fallback;
      },
      credentialName: 'm365Agent365Api',
      credentials: {
        tenantId: 't', blueprintAppId: 'bp', transport: 'sidecar',
        sidecarUrl: 'http://s:5000', outboundDownstreamApi: 'MessagingBotApi',
        validateVia: 'sameAsOutbound',
      },
    });
    await updateExecute.call(ctx, 0, 'agent365' as any, {
      tenantId: 't', blueprintAppId: 'bp', transport: 'sidecar',
      sidecarUrl: 'http://s:5000', outboundDownstreamApi: 'MessagingBotApi',
      validateVia: 'sameAsOutbound',
    } as any, new Map());
    expect(acquireOutboundToken).toHaveBeenCalledWith(expect.objectContaining({
      identityMode: 'agentUser', agentUsername: 'a@b.com',
    }));
  });
});

// ── deleteMessage ─────────────────────────────────────────────────────────────

describe('message/deleteMessage routing', () => {
  beforeEach(() => (acquireOutboundToken as any).mockClear());

  it('classic authKind → router receives classicBot', async () => {
    const ctx = makeExecuteMock({
      items: ENVELOPE_INPUT,
      nodeParams: (name: string, _i: number, fallback?: unknown) => {
        const map: Record<string, unknown> = {
          authKind: 'classicBot',
          resource: 'message', operation: 'delete',
          conversationSource: 'envelope',
        };
        return map[name] ?? fallback;
      },
      credentialName: 'm365AgentApi',
      credentials: { appType: 'SingleTenant', clientId: 'c', tenantId: 't', clientSecret: 's' },
    });
    const bundles = new Map<string, BotConnectorBundle>();
    await deleteExecute.call(ctx, 0, 'classicBot' as any, { appType: 'SingleTenant', clientId: 'c' } as any, bundles);
    expect(acquireOutboundToken).toHaveBeenCalledWith(expect.objectContaining({ authKind: 'classicBot' }));
  });

  it('agent365 + autonomous → router receives downstreamApi=MessagingBotApi', async () => {
    const ctx = makeExecuteMock({
      items: ENVELOPE_INPUT,
      nodeParams: (name: string, _i: number, fallback?: unknown) => {
        const map: Record<string, unknown> = {
          authKind: 'agent365', identityMode: 'autonomous',
          resource: 'message', operation: 'delete',
          conversationSource: 'envelope',
        };
        return map[name] ?? fallback;
      },
      credentialName: 'm365Agent365Api',
      credentials: {
        tenantId: 't', blueprintAppId: 'bp', transport: 'sidecar',
        sidecarUrl: 'http://s:5000', outboundDownstreamApi: 'MessagingBotApi',
        validateVia: 'sameAsOutbound',
      },
    });
    await deleteExecute.call(ctx, 0, 'agent365' as any, {
      tenantId: 't', blueprintAppId: 'bp', transport: 'sidecar',
      sidecarUrl: 'http://s:5000', outboundDownstreamApi: 'MessagingBotApi',
      validateVia: 'sameAsOutbound',
    } as any, new Map());
    expect(acquireOutboundToken).toHaveBeenCalledWith(expect.objectContaining({
      authKind: 'agent365', identityMode: 'autonomous', downstreamApi: 'MessagingBotApi',
    }));
  });

  it('agent365 + agentUser by UPN → forwards agentUsername', async () => {
    const ctx = makeExecuteMock({
      items: ENVELOPE_INPUT,
      nodeParams: (name: string, _i: number, fallback?: unknown) => {
        const map: Record<string, unknown> = {
          authKind: 'agent365', identityMode: 'agentUser',
          userSelectorMode: 'byUpn', agentUsername: 'a@b.com',
          resource: 'message', operation: 'delete',
          conversationSource: 'envelope',
        };
        return map[name] ?? fallback;
      },
      credentialName: 'm365Agent365Api',
      credentials: {
        tenantId: 't', blueprintAppId: 'bp', transport: 'sidecar',
        sidecarUrl: 'http://s:5000', outboundDownstreamApi: 'MessagingBotApi',
        validateVia: 'sameAsOutbound',
      },
    });
    await deleteExecute.call(ctx, 0, 'agent365' as any, {
      tenantId: 't', blueprintAppId: 'bp', transport: 'sidecar',
      sidecarUrl: 'http://s:5000', outboundDownstreamApi: 'MessagingBotApi',
      validateVia: 'sameAsOutbound',
    } as any, new Map());
    expect(acquireOutboundToken).toHaveBeenCalledWith(expect.objectContaining({
      identityMode: 'agentUser', agentUsername: 'a@b.com',
    }));
  });
});

// ── replyInThread ─────────────────────────────────────────────────────────────

describe('message/replyInThread routing', () => {
  beforeEach(() => (acquireOutboundToken as any).mockClear());

  it('classic authKind → router receives classicBot', async () => {
    const ctx = makeExecuteMock({
      items: ENVELOPE_INPUT,
      nodeParams: (name: string, _i: number, fallback?: unknown) => {
        const map: Record<string, unknown> = {
          authKind: 'classicBot',
          resource: 'message', operation: 'replyInThread',
          conversationSource: 'envelope',
          parentActivityId: 'root-1',
          text: 'hi', options: {},
        };
        return map[name] ?? fallback;
      },
      credentialName: 'm365AgentApi',
      credentials: { appType: 'SingleTenant', clientId: 'c', tenantId: 't', clientSecret: 's' },
    });
    const bundles = new Map<string, BotConnectorBundle>();
    await replyInThreadExecute.call(ctx, 0, 'classicBot' as any, { appType: 'SingleTenant', clientId: 'c' } as any, bundles);
    expect(acquireOutboundToken).toHaveBeenCalledWith(expect.objectContaining({ authKind: 'classicBot' }));
  });

  it('agent365 + autonomous → router receives downstreamApi=MessagingBotApi', async () => {
    const ctx = makeExecuteMock({
      items: ENVELOPE_INPUT,
      nodeParams: (name: string, _i: number, fallback?: unknown) => {
        const map: Record<string, unknown> = {
          authKind: 'agent365', identityMode: 'autonomous',
          resource: 'message', operation: 'replyInThread',
          conversationSource: 'envelope',
          parentActivityId: 'root-1',
          text: 'hi', options: {},
        };
        return map[name] ?? fallback;
      },
      credentialName: 'm365Agent365Api',
      credentials: {
        tenantId: 't', blueprintAppId: 'bp', transport: 'sidecar',
        sidecarUrl: 'http://s:5000', outboundDownstreamApi: 'MessagingBotApi',
        validateVia: 'sameAsOutbound',
      },
    });
    await replyInThreadExecute.call(ctx, 0, 'agent365' as any, {
      tenantId: 't', blueprintAppId: 'bp', transport: 'sidecar',
      sidecarUrl: 'http://s:5000', outboundDownstreamApi: 'MessagingBotApi',
      validateVia: 'sameAsOutbound',
    } as any, new Map());
    expect(acquireOutboundToken).toHaveBeenCalledWith(expect.objectContaining({
      authKind: 'agent365', identityMode: 'autonomous', downstreamApi: 'MessagingBotApi',
    }));
  });

  it('agent365 + agentUser by UPN → forwards agentUsername', async () => {
    const ctx = makeExecuteMock({
      items: ENVELOPE_INPUT,
      nodeParams: (name: string, _i: number, fallback?: unknown) => {
        const map: Record<string, unknown> = {
          authKind: 'agent365', identityMode: 'agentUser',
          userSelectorMode: 'byUpn', agentUsername: 'a@b.com',
          resource: 'message', operation: 'replyInThread',
          conversationSource: 'envelope',
          parentActivityId: 'root-1',
          text: 'hi', options: {},
        };
        return map[name] ?? fallback;
      },
      credentialName: 'm365Agent365Api',
      credentials: {
        tenantId: 't', blueprintAppId: 'bp', transport: 'sidecar',
        sidecarUrl: 'http://s:5000', outboundDownstreamApi: 'MessagingBotApi',
        validateVia: 'sameAsOutbound',
      },
    });
    await replyInThreadExecute.call(ctx, 0, 'agent365' as any, {
      tenantId: 't', blueprintAppId: 'bp', transport: 'sidecar',
      sidecarUrl: 'http://s:5000', outboundDownstreamApi: 'MessagingBotApi',
      validateVia: 'sameAsOutbound',
    } as any, new Map());
    expect(acquireOutboundToken).toHaveBeenCalledWith(expect.objectContaining({
      identityMode: 'agentUser', agentUsername: 'a@b.com',
    }));
  });
});
