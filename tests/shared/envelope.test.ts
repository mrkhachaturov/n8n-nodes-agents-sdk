import { describe, it, expect } from 'vitest';
import {
	activityToConversationReference,
	parseActivity,
	mergeEnvelope,
} from '../../shared/envelope';

const incoming = {
	type: 'message',
	id: 'activity-123',
	timestamp: '2026-04-16T12:00:00Z',
	serviceUrl: 'https://smba.trafficmanager.net/emea/',
	channelId: 'msteams',
	from: { id: 'u-1', name: 'Alice', aadObjectId: 'aad-1', role: 'user' },
	conversation: { id: 'c-1', conversationType: 'channel' },
	recipient: { id: 'bot-1', name: 'Bot' },
	locale: 'en-US',
	text: 'hello',
	value: { action: 'accept', order_number: 'ORD-42' },
	replyToId: 'parent-999',
};

describe('shared/envelope', () => {
	it('activityToConversationReference extracts routing fields', () => {
		const ref = activityToConversationReference(incoming as any);
		expect(ref.serviceUrl).toBe('https://smba.trafficmanager.net/emea/');
		expect(ref.conversation.id).toBe('c-1');
		expect(ref.activityId).toBe('activity-123');
		expect(ref.channelId).toBe('msteams');
		expect(ref.user?.id).toBe('u-1');
		expect(ref.user?.aadObjectId).toBe('aad-1');
	});

	it('parseActivity flattens common action-submit fields', () => {
		const parsed = parseActivity(incoming as any);
		expect(parsed.type).toBe('message');
		expect(parsed.text).toBe('hello');
		expect(parsed.action).toBe('accept');
		expect(parsed.submitData).toEqual({ action: 'accept', order_number: 'ORD-42' });
		expect(parsed.userName).toBe('Alice');
		expect(parsed.userId).toBe('u-1');
		expect(parsed.aadObjectId).toBe('aad-1');
	});

	it('mergeEnvelope preserves conversationReference when builder replaces activity (reply path)', () => {
		const input = {
			conversationReference: {
				serviceUrl: 'x',
				conversation: { id: 'c' },
				activityId: 'a',
				channelId: 'msteams',
			},
			activity: { type: 'message', text: 'old' },
		};
		const newActivity = { type: 'message', text: 'new' };
		const merged = mergeEnvelope(input, newActivity);
		expect(merged.conversationReference).toBe(input.conversationReference);
		expect(merged.activity).toEqual(newActivity);
	});

	it('mergeEnvelope emits activity-only when conversationReference is absent (proactive path)', () => {
		const newActivity = { type: 'message', text: 'proactive hello' };
		const merged = mergeEnvelope({ activity: { type: 'message' } }, newActivity);
		expect(merged.conversationReference).toBeUndefined();
		expect(merged.activity).toEqual(newActivity);
	});
});

describe('parseActivity — 0.5.0 additions', () => {
	it('populates membersAdded/Removed on conversationUpdate', () => {
		const parsed = parseActivity({
			type: 'conversationUpdate',
			membersAdded: [{ id: 'user1', name: 'Alice' }],
			membersRemoved: [{ id: 'user2', name: 'Bob' }],
		} as any);
		expect(parsed.membersAdded).toEqual([{ id: 'user1', name: 'Alice' }]);
		expect(parsed.membersRemoved).toEqual([{ id: 'user2', name: 'Bob' }]);
	});

	it('populates reactionsAdded/Removed on messageReaction', () => {
		const parsed = parseActivity({
			type: 'messageReaction',
			reactionsAdded: [{ type: 'like', user: { id: 'user1', name: 'Alice' } }],
			reactionsRemoved: [{ type: 'heart', user: { id: 'user2', name: 'Bob' } }],
		} as any);
		expect(parsed.reactionsAdded).toEqual([{ type: 'like', user: { id: 'user1', name: 'Alice' } }]);
		expect(parsed.reactionsRemoved).toEqual([
			{ type: 'heart', user: { id: 'user2', name: 'Bob' } },
		]);
	});

	it('leaves parsed sub-fields undefined when activity type does not match', () => {
		const parsed = parseActivity({ type: 'message', text: 'hello' } as any);
		expect(parsed.membersAdded).toBeUndefined();
		expect(parsed.reactionsAdded).toBeUndefined();
	});

	it('does NOT assign invokeName on parsed (envelope-top-level per spec)', () => {
		const parsed = parseActivity({ type: 'invoke', name: 'taskModule/fetch' } as any);
		expect((parsed as any).invokeName).toBeUndefined();
	});
});
