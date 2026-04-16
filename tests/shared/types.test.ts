import { describe, it, expect } from 'vitest';
import type { ConversationReference, ItemEnvelope, Operation } from '../../shared/types';

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

	it('Operation type covers all five operations', () => {
		const ops: Operation[] = ['reply', 'proactive', 'update', 'delete', 'replyInThread'];
		expect(ops).toHaveLength(5);
	});
});
