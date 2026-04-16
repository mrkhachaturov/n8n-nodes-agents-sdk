/**
 * Structural alias for an Activity-compatible object.
 * Uses a structural type rather than importing from @microsoft/agents-activity
 * to keep this file free of restricted imports per n8n community node rules.
 * Nodes that need the full SDK Activity type import it directly from the SDK.
 */
export type ActivityLike = {
	type: string;
	id?: string;
	text?: string;
	channelId?: string;
	attachments?: unknown[];
	entities?: unknown[];
	value?: unknown;
	replyToId?: string;
	[key: string]: unknown;
};

/**
 * Minimal ConversationReference shape used across the package.
 * Mirrors the SDK's ConversationReference but restricts it to the fields
 * the Trigger always emits and the sender always consumes.
 */
export interface ConversationReference {
	serviceUrl: string;
	conversation: { id: string; conversationType?: string; isGroup?: boolean };
	activityId: string;
	bot?: { id?: string; name?: string };
	user?: { id?: string; name?: string; aadObjectId?: string; role?: string };
	channelId: string;
	locale?: string;
}

/**
 * Parsed convenience fields extracted from the incoming Activity for
 * common downstream use cases (e.g. Action.Submit payloads).
 */
export interface ParsedActivity {
	type: string;
	text?: string;
	action?: string;
	submitData?: Record<string, unknown>;
	userName?: string;
	userId?: string;
	aadObjectId?: string;
	timestamp?: string;
}

/**
 * The shape every node in this package reads from and writes to.
 * See design spec, section "Item envelope convention".
 *
 * conversationReference is OPTIONAL — builder nodes may emit activity-only
 * items when used on proactive paths originating from non-bot webhooks
 * (e.g., a 1C ERP event). In that case M365SendActivity proactive receives
 * the target conversationReference as explicit node input.
 */
export interface ItemEnvelope {
	conversationReference?: ConversationReference;
	activity: Partial<ActivityLike>;
	parsed?: ParsedActivity;
	/** Full incoming Activity from the SDK — typed as unknown to avoid restricted imports. */
	raw?: unknown;
}

/** Operations supported by M365SendActivity. */
export type Operation = 'reply' | 'proactive' | 'update' | 'delete' | 'replyInThread';

/** Azure Bot Service app type selection on the credential. */
export type AppType = 'SingleTenant' | 'MultiTenant' | 'UserAssignedMsi';

/** Fields collected by the M365AgentApi credential. */
export interface M365AgentCredentials {
	appType: AppType;
	clientId: string;
	clientSecret?: string;
	tenantId?: string;
	anonymousAllowed?: boolean;
}
