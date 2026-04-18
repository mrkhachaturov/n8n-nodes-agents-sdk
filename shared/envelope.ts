import type { Activity } from '@microsoft/agents-activity';
import type { ConversationReference, ItemEnvelope, ParsedActivity } from './types';

/**
 * Extract routing fields from a received Activity into a ConversationReference.
 * Sender nodes use this to know where to POST without needing the raw Activity.
 */
export function activityToConversationReference(activity: Activity): ConversationReference {
	if (!activity.serviceUrl) throw new Error('activity.serviceUrl missing');
	if (!activity.conversation?.id) throw new Error('activity.conversation.id missing');
	if (!activity.id) throw new Error('activity.id missing');

	return {
		serviceUrl: activity.serviceUrl,
		conversation: {
			id: activity.conversation.id,
			conversationType: activity.conversation.conversationType,
			isGroup: activity.conversation.isGroup,
		},
		activityId: activity.id,
		channelId: activity.channelId ?? 'unknown',
		bot: activity.recipient
			? { id: activity.recipient.id, name: activity.recipient.name }
			: undefined,
		user: activity.from
			? {
					id: activity.from.id,
					name: activity.from.name,
					// eslint-disable-next-line @typescript-eslint/no-explicit-any
					aadObjectId: (activity.from as any).aadObjectId,
					role: activity.from.role as string | undefined,
				}
			: undefined,
		// eslint-disable-next-line @typescript-eslint/no-explicit-any
		locale: (activity as any).locale,
	};
}

/**
 * Derive a flat ParsedActivity from the incoming Activity.
 * Typical use: downstream workflow needs `action` from a card's Action.Submit
 * without reaching into `activity.value`.
 */
export function parseActivity(activity: Activity): ParsedActivity {
	// eslint-disable-next-line @typescript-eslint/no-explicit-any
	const value = (activity as any).value;
	return {
		type: activity.type ?? 'unknown',
		text: activity.text,
		// eslint-disable-next-line @typescript-eslint/no-explicit-any
		action: typeof value === 'object' && value !== null ? (value as any).action : undefined,
		submitData:
			typeof value === 'object' && value !== null ? (value as Record<string, unknown>) : undefined,
		userName: activity.from?.name,
		userId: activity.from?.id,
		// eslint-disable-next-line @typescript-eslint/no-explicit-any
		aadObjectId: (activity.from as any)?.aadObjectId,
		timestamp:
			activity.timestamp instanceof Date
				? activity.timestamp.toISOString()
				: (activity.timestamp as string | undefined),
	};
}

/**
 * Classify the inbound bearer token source from its validated claims.
 * Used by the trigger to populate AuthContext.tokenSource for agent365 flows.
 */
export function detectTokenSource(
	claims: Record<string, unknown>,
): 'botFramework' | 'messagingBotApi' | 'agenticIdentity' | 'unknown' {
	const aud = String(claims.aud ?? '');
	if (aud.includes('botframework')) return 'botFramework';
	if (aud.includes('a6c6ce43')) return 'messagingBotApi';
	if (aud.startsWith('api://')) return 'agenticIdentity';
	return 'unknown';
}

/**
 * Merge a new Activity into an envelope.
 * Preserves conversationReference if present (reply-style path);
 * otherwise returns activity-only (proactive-from-non-bot path).
 * Builder nodes MUST use this — never spread over the full item.
 */
export function mergeEnvelope(
	current: { conversationReference?: ConversationReference; activity?: Partial<Activity> },
	nextActivity: Partial<Activity>,
): ItemEnvelope {
	if (current.conversationReference) {
		return { conversationReference: current.conversationReference, activity: nextActivity };
	}
	return { activity: nextActivity };
}
