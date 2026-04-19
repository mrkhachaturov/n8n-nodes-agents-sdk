import { describe, it, expect } from 'vitest';
import {
	serializeConversationRef,
	deserializeConversationRef,
	redactBearerInError,
} from '../../shared/conversationRefSerializer';

const envelopeWithAgent365 = {
	conversationReference: {
		serviceUrl: 'https://test.botservice',
		conversation: { id: 'conv1' },
		bot: { id: 'b' },
		user: { id: 'u' },
		channelId: 'msteams',
		locale: 'en',
		activityId: 'act-1',
	},
	authContext: {
		inboundBearer: 'Bearer eyJabc.xyz',
		tokenSource: 'agent365',
		validatedClaims: { aud: 'blueprint-id', appId: 'app-id', tid: 'tenant', exp: 9999999999 },
	},
};

describe('conversationRefSerializer', () => {
	describe('serialize', () => {
		it('bearer off by default: excludes inboundBearer from payload', () => {
			const payload = serializeConversationRef(envelopeWithAgent365 as any, {
				includeBearer: false,
			});
			expect(payload.conversationReference).toBeDefined();
			expect(payload.tokenSource).toBe('agent365');
			expect(payload.claimsExp).toBe(9999999999);
			expect((payload as any).inboundBearer).toBeUndefined();
		});

		it('bearer on: includes inboundBearer in payload', () => {
			const payload = serializeConversationRef(envelopeWithAgent365 as any, {
				includeBearer: true,
			});
			expect(payload.inboundBearer).toBe('Bearer eyJabc.xyz');
		});

		it('rejects save when envelope is missing conversationReference', () => {
			expect(() =>
				serializeConversationRef({} as any, { includeBearer: false }),
			).toThrow(/conversationReference/);
		});

		it('works when envelope has no authContext (classic bot)', () => {
			const classicEnvelope = {
				conversationReference: envelopeWithAgent365.conversationReference,
			};
			const payload = serializeConversationRef(classicEnvelope as any, {
				includeBearer: false,
			});
			expect(payload.conversationReference).toEqual(envelopeWithAgent365.conversationReference);
			expect(payload.tokenSource).toBeUndefined();
			expect(payload.claimsExp).toBeUndefined();
		});
	});

	describe('deserialize', () => {
		it('reconstructs envelope shape from a saved payload', () => {
			const payload = serializeConversationRef(envelopeWithAgent365 as any, {
				includeBearer: true,
			});
			const result = deserializeConversationRef(payload);
			expect(result.conversationReference).toEqual(envelopeWithAgent365.conversationReference);
			expect(result.authContext?.inboundBearer).toBe('Bearer eyJabc.xyz');
			expect(result.authContext?.tokenSource).toBe('agent365');
			expect(result.expired).toBe(false);
		});

		it('flags expired=true when claimsExp is past now', () => {
			const expiredPayload = {
				...serializeConversationRef(envelopeWithAgent365 as any, { includeBearer: true }),
				claimsExp: 1, // January 1970
			};
			const result = deserializeConversationRef(expiredPayload);
			expect(result.expired).toBe(true);
			expect(result.authContext?.inboundBearer).toBe('Bearer eyJabc.xyz'); // still returned
		});

		it('throws on malformed payload (missing conversationReference)', () => {
			expect(() => deserializeConversationRef({} as any)).toThrow(/conversationReference.*missing/i);
		});
	});

	describe('redactBearerInError', () => {
		it('replaces bearer token with [REDACTED] in error messages', () => {
			const err = new Error('Failed to save: inboundBearer=Bearer eyJabc.xyz something else');
			const redacted = redactBearerInError(err);
			expect(redacted.message).not.toContain('eyJabc.xyz');
			expect(redacted.message).toContain('[REDACTED]');
		});
	});
});
