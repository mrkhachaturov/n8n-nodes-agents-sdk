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

function loadCtx(payload: any, payloadField = 'payload') {
	return {
		getInputData: () => [{ json: { [payloadField]: payload } }],
		getNodeParameter: (n: string) => {
			if (n === 'resource') return 'serialize';
			if (n === 'operation') return 'load';
			if (n === 'payloadField') return payloadField;
			return undefined;
		},
		getNode: () => ({ name: 'test' }),
		helpers: { returnJsonArray: (x: any) => x },
	} as any;
}

describe('M365ConversationRef — Serialize/load', () => {
	const node = new M365ConversationRef();

	const savedPayload = {
		conversationReference: { serviceUrl: 'u', conversation: { id: 'c' } },
		tokenSource: 'agent365',
		claimsExp: 9999999999,
		inboundBearer: 'Bearer eyJ',
	};

	it('reconstructs envelope shape with expired=false', async () => {
		const res = await node.execute.call(loadCtx(savedPayload) as any);
		const out = res[0][0].json as any;
		expect(out.conversationReference).toEqual(savedPayload.conversationReference);
		expect(out.authContext.inboundBearer).toBe('Bearer eyJ');
		expect(out.expired).toBe(false);
	});

	it('reports expired=true when claimsExp is in the past', async () => {
		const expired = { ...savedPayload, claimsExp: 1 };
		const res = await node.execute.call(loadCtx(expired) as any);
		const out = res[0][0].json as any;
		expect(out.expired).toBe(true);
		// Bearer still returned — workflow decides how to react
		expect(out.authContext.inboundBearer).toBe('Bearer eyJ');
	});

	it('respects a custom payloadField', async () => {
		const res = await node.execute.call(loadCtx(savedPayload, 'stored') as any);
		const out = res[0][0].json as any;
		expect(out.conversationReference).toEqual(savedPayload.conversationReference);
	});

	it('throws on malformed payload (missing conversationReference)', async () => {
		await expect(node.execute.call(loadCtx({ tokenSource: 'agent365' }) as any)).rejects.toThrow(/conversationReference/);
	});
});
