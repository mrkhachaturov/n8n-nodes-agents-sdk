import type { Activity } from '@microsoft/agents-activity';

/**
 * Minimal ConversationReference shape used across the package.
 * Mirrors the SDK's ConversationReference but restricts it to the fields
 * the Trigger emits and the sender consumes.
 *
 * activityId is OPTIONAL: the Trigger always populates it (see
 * activityToConversationReference), but user-crafted proactive targets
 * have no inbound activity to reference. The M365 Agent node enforces
 * per-operation presence at runtime — reply / update / delete throw
 * NodeOperationError when activityId is missing.
 */
export interface ConversationReference {
	serviceUrl: string;
	conversation: { id: string; conversationType?: string; isGroup?: boolean };
	activityId?: string;
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
 *
 * conversationReference is OPTIONAL — proactive sends originating from a
 * non-bot webhook (a schedule, an external event, a database row) arrive
 * without an inbound envelope, and the M365 Agent node's "Specify Manually"
 * Conversation Source builds the reference from UI fields at execute time.
 */
export interface ItemEnvelope {
	conversationReference?: ConversationReference;
	activity: Partial<Activity>;
	parsed?: ParsedActivity;
	raw?: Activity;
}

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

// === Agent 365 additions (v0.3.0) ===

/** Top-level node parameter that picks the credential family. */
export type AuthKind = 'classicBot' | 'agent365';

/** Per-outbound-operation identity selector for Agent 365 credentials. */
export type IdentityMode = 'autonomous' | 'agentUser' | 'interactiveOBO';

/** Inline vs sidecar token-acquisition transport for Agent 365. */
export type M365Agent365Transport = 'inline' | 'sidecar';

/** Inline-only credential kind. */
export type M365Agent365InlineCredKind = 'clientSecret' | 'clientCert';

/** Validator routing when authKind=agent365. */
export type M365Agent365ValidateVia = 'sameAsOutbound' | 'inline' | 'sidecar';

/** Classic bot credential shape — alias of M365AgentCredentials for router typing clarity. */
export type M365ClassicBotCred = M365AgentCredentials;

/** Agent 365 credential shape (new in v0.3.0). */
export interface M365Agent365Cred {
	tenantId: string;
	blueprintAppId: string;
	/**
	 * Agent Identity instance appId (distinct from blueprintAppId).
	 * Used as the `AgentIdentity` query param on sidecar calls.
	 * Optional here for backwards compat with early fixtures; required at
	 * runtime — backend falls back to blueprintAppId if unset.
	 */
	agentInstanceAppId?: string;
	transport: M365Agent365Transport;
	inlineCredKind?: M365Agent365InlineCredKind;
	blueprintSecret?: string;
	blueprintCertPem?: string;
	blueprintCertThumbprint?: string;
	sidecarUrl?: string;
	outboundDownstreamApi?: string;
	defaultAgentUsername?: string;
	validateVia: M365Agent365ValidateVia;
	validatorSidecarUrl?: string;
	options?: {
		allowedIssuers?: string;
		allowedAudiences?: string;
		sidecarTimeoutMs?: number;
		tokenCacheTtlSec?: number;
	};
}
