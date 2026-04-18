import { describe, it, expect } from 'vitest';
import type {
  ConversationReference,
  ItemEnvelope,
  AuthKind,
  IdentityMode,
  M365ClassicBotCred,
  M365Agent365Cred,
  M365Agent365Transport,
} from '../../shared/types';

describe('shared/types', () => {
	it('ConversationReference has required routing fields', () => {
		const ref: ConversationReference = {
			serviceUrl: 'https://smba.trafficmanager.net/emea/',
			conversation: { id: '19:x@thread.tacv2', conversationType: 'channel' },
			activityId: '1234',
			bot: { id: 'bot-id', name: 'Bot' },
			user: { id: 'u', name: 'U' },
			channelId: 'msteams',
			locale: 'en-US',
		};
		expect(ref.serviceUrl).toBeDefined();
		expect(ref.conversation.id).toBeDefined();
		expect(ref.activityId).toBeDefined();
	});

	it('ItemEnvelope combines conversationReference and activity (reply-style)', () => {
		const env: ItemEnvelope = {
			conversationReference: {
				serviceUrl: 'x',
				conversation: { id: 'c' },
				activityId: 'a',
				channelId: 'msteams',
			},
			activity: { type: 'message', text: 'hi' },
		};
		expect(env.conversationReference).toBeDefined();
		expect(env.activity).toBeDefined();
	});

	it('ItemEnvelope can carry activity only (proactive-path builders)', () => {
		const env: ItemEnvelope = {
			activity: { type: 'message', text: 'hi' },
		};
		expect(env.conversationReference).toBeUndefined();
		expect(env.activity).toBeDefined();
	});

});

describe('auth types', () => {
	it('AuthKind admits exactly two values', () => {
		const valid: AuthKind[] = ['classicBot', 'agent365'];
		expect(valid).toHaveLength(2);
	});

	it('IdentityMode admits three values', () => {
		const valid: IdentityMode[] = ['autonomous', 'agentUser', 'interactiveOBO'];
		expect(valid).toHaveLength(3);
	});

	it('M365Agent365Cred has required agent365 shape', () => {
		const cred: M365Agent365Cred = {
			tenantId: 't',
			blueprintAppId: 'b',
			transport: 'sidecar',
			sidecarUrl: 'http://x:5000',
			outboundDownstreamApi: 'MessagingBotApi',
			validateVia: 'sameAsOutbound',
		};
		expect(cred.blueprintAppId).toBe('b');
	});

	it('M365Agent365Transport admits inline or sidecar', () => {
		const t: M365Agent365Transport[] = ['inline', 'sidecar'];
		expect(t).toHaveLength(2);
	});
});
