import type { ItemEnvelope, AuthContext, ConversationReference } from './types';

/**
 * Payload shape persisted by the M365ConversationRef save path and consumed by
 * its load path. Kept narrow on purpose: only the routing reference plus a
 * minimal, opt-in auth summary. Bearers are excluded by default and MUST only
 * be persisted when the caller explicitly sets `includeBearer: true`.
 *
 * Capability: T3-conversation-reference-persistence (foundation).
 */
export interface SerializedConversationRef {
	conversationReference: ConversationReference;
	tokenSource?: string;
	claimsExp?: number;
	/** Opt-in only. Present iff options.includeBearer was true at serialize time. */
	inboundBearer?: string;
}

/**
 * Result of deserializing a saved payload. `expired` is derived from
 * `claimsExp` vs current wall-clock; callers decide how to react (warn, refuse
 * to send, or accept stale references for read-only flows).
 */
export interface DeserializedResult {
	conversationReference: ConversationReference;
	authContext?: AuthContext;
	expired: boolean;
}

/**
 * Serialize an ItemEnvelope into the narrow persistence payload. Throws when
 * the envelope has no conversationReference — save paths need an explicit
 * target, never an implicit one.
 */
export function serializeConversationRef(
	envelope: ItemEnvelope,
	options: { includeBearer: boolean },
): SerializedConversationRef {
	if (!envelope || !envelope.conversationReference) {
		throw new Error(
			'Serialize/save requires envelope.conversationReference. Configure Conversation Source: From Envelope, or supply conversationReference explicitly.',
		);
	}
	const payload: SerializedConversationRef = {
		conversationReference: envelope.conversationReference,
	};
	if (envelope.authContext) {
		payload.tokenSource = envelope.authContext.tokenSource as string;
		const expRaw = envelope.authContext.validatedClaims?.exp;
		if (typeof expRaw === 'number') payload.claimsExp = expRaw;
		if (options.includeBearer) {
			payload.inboundBearer = envelope.authContext.inboundBearer;
		}
	}
	return payload;
}

/**
 * Reconstruct an envelope-shaped result from a saved payload. Throws when the
 * payload is missing conversationReference — a malformed record cannot drive
 * any outbound operation.
 */
export function deserializeConversationRef(
	payload: SerializedConversationRef,
): DeserializedResult {
	if (!payload || !payload.conversationReference) {
		throw new Error(
			'Serialize/load: payload.conversationReference is missing. Verify the save path preserved the payload shape.',
		);
	}
	const now = Math.floor(Date.now() / 1000);
	const expired = typeof payload.claimsExp === 'number' && payload.claimsExp < now;
	const result: DeserializedResult = {
		conversationReference: payload.conversationReference,
		expired,
	};
	if (payload.tokenSource || payload.inboundBearer || payload.claimsExp) {
		result.authContext = {
			inboundBearer: payload.inboundBearer ?? '',
			tokenSource: payload.tokenSource as AuthContext['tokenSource'],
			validatedClaims: payload.claimsExp ? { exp: payload.claimsExp } : {},
		};
	}
	return result;
}

const BEARER_PATTERN = /Bearer\s+[A-Za-z0-9._-]+/g;

/**
 * Return a copy of the given Error with any `Bearer <token>` fragments in
 * `message` and `stack` replaced by `[REDACTED]`. Used by save/load error
 * paths so that a thrown error never leaks the inbound bearer token into n8n
 * execution logs.
 */
export function redactBearerInError(err: Error): Error {
	const redacted = new Error(err.message.replace(BEARER_PATTERN, '[REDACTED]'));
	redacted.name = err.name;
	redacted.stack = err.stack?.replace(BEARER_PATTERN, '[REDACTED]');
	return redacted;
}
