import axios, { type AxiosInstance } from 'axios';
import {
	ConnectorClient,
	MsalTokenProvider,
	type AuthConfiguration,
} from '@microsoft/agents-hosting';
import type { Activity } from '@microsoft/agents-activity';

const BOT_FRAMEWORK_SCOPE = 'https://api.botframework.com';

export interface BotConnectorBundle {
	/** SDK client for reply/proactive/update/delete. */
	client: ConnectorClient;
	/** Raw axios instance for operations the SDK does not cover (replyInThread). */
	axios: AxiosInstance;
	/** Bearer token acquired at bundle creation time. */
	token: string;
	/** Base URL the connector was created against. */
	baseURL: string;
}

/**
 * Build a BotConnectorBundle from a pre-acquired Authorization header.
 * Used by operations that go through the shared auth router (acquireOutboundToken),
 * which handles both classicBot (MSAL) and agent365 (sidecar / inline) paths.
 * The header should be in "Bearer <token>" form; a bare token is also accepted.
 */
export function createConnectorFromBearer(
	serviceUrl: string,
	authorizationHeader: string,
): BotConnectorBundle {
	const baseURL = serviceUrl.endsWith('/') ? serviceUrl : `${serviceUrl}/`;
	const token = authorizationHeader.startsWith('Bearer ')
		? authorizationHeader.slice('Bearer '.length)
		: authorizationHeader;
	const client = ConnectorClient.createClientWithToken(baseURL, token);
	const ax = axios.create({
		baseURL,
		headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
	});
	return { client, axios: ax, token, baseURL };
}

/**
 * Build a BotConnectorBundle for outbound Bot Connector calls.
 * Acquires a token via the SDK's MSAL provider and wraps both the SDK client
 * and a raw axios instance sharing the same bearer token.
 *
 * @deprecated Use createConnectorFromBearer with acquireOutboundToken instead.
 *   This function will be removed in Task 4.8 once all operations migrate to
 *   the shared auth router.
 */
export async function createConnector(
	authConfig: AuthConfiguration,
	serviceUrl: string,
): Promise<BotConnectorBundle> {
	const baseURL = serviceUrl.endsWith('/') ? serviceUrl : `${serviceUrl}/`;
	const provider = new MsalTokenProvider();
	const token = await provider.getAccessToken(authConfig, BOT_FRAMEWORK_SCOPE);

	const client = ConnectorClient.createClientWithToken(baseURL, token);
	const ax = axios.create({
		baseURL,
		headers: {
			Authorization: `Bearer ${token}`,
			'Content-Type': 'application/json',
		},
	});

	return { client, axios: ax, token, baseURL };
}

/**
 * Encode a conversation ID for use in a Bot Connector URL path segment.
 * Teams conversation IDs use `:` and `@` which are valid in URI path segments
 * and must not be percent-encoded (the Bot Connector rejects encoded forms).
 * Only truly ambiguous path delimiters (`/`, `?`) are encoded.
 */
function encodeConversationId(id: string): string {
	return encodeURIComponent(id).replace(/%3A/g, ':').replace(/%40/g, '@');
}

/**
 * Build the relative URL for the Teams-specific thread reply. The SDK's
 * ConnectorClient does not expose this primitive because it is a URL trick,
 * not a new protocol endpoint.
 */
export function buildReplyInThreadUrl(conversationId: string, parentActivityId: string): string {
	return `v3/conversations/${encodeConversationId(conversationId)};messageid=${encodeURIComponent(parentActivityId)}/activities`;
}

/**
 * Post a reply-in-thread Activity via the raw axios instance.
 */
export async function replyInThread(
	bundle: BotConnectorBundle,
	conversationId: string,
	parentActivityId: string,
	activity: Partial<Activity>,
): Promise<{ id?: string }> {
	const url = buildReplyInThreadUrl(conversationId, parentActivityId);
	const response = await bundle.axios.post<{ id?: string }>(url, activity);
	return response.data;
}
