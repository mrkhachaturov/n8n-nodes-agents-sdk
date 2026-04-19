import { describe, it, expect } from 'vitest';
import { M365ConversationRef } from '../../../nodes/M365ConversationRef/M365ConversationRef.node';

function ctx(envelope: any, includeBearer = false) {
	return {
		getInputData: () => [{ json: envelope }],
		getNodeParameter: (n: string) => {
			if (n === 'resource') return 'serialize';
			if (n === 'operation') return 'save';
			if (n === 'includeBearer') return includeBearer;
			return undefined;
		},
		getNode: () => ({ name: 'test' }),
		helpers: { returnJsonArray: (x: any) => x },
	} as any;
}

describe('M365ConversationRef — Serialize/save', () => {
	const node = new M365ConversationRef();
	const envelope = {
		conversationReference: { serviceUrl: 'u', conversation: { id: 'c' } },
		authContext: {
			inboundBearer: 'Bearer eyJ',
			// NOTE: 'agent365' is not in AuthContext.tokenSource today (union is
			// botFramework | messagingBotApi | agenticIdentity | unknown). Cast
			// keeps the fixture aligned with the plan's future-facing intent.
			tokenSource: 'agent365' as any,
			validatedClaims: { exp: 9999999999 },
		},
	};

	it('emits payload with bearer off by default', async () => {
		const res = await node.execute.call(ctx(envelope, false) as any);
		const payload = res[0][0].json as any;
		expect(payload.conversationReference).toEqual(envelope.conversationReference);
		expect(payload.tokenSource).toBe('agent365');
		expect(payload.claimsExp).toBe(9999999999);
		expect(payload.inboundBearer).toBeUndefined();
	});

	it('emits payload with bearer when toggle on', async () => {
		const res = await node.execute.call(ctx(envelope, true) as any);
		const payload = res[0][0].json as any;
		expect(payload.inboundBearer).toBe('Bearer eyJ');
	});

	it('throws on missing envelope', async () => {
		await expect(node.execute.call(ctx({}, false) as any)).rejects.toThrow(/conversationReference/);
	});
});
