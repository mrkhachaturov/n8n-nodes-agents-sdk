import type { M365Agent365Cred } from '../types';
import type { AuthConfiguration } from '@microsoft/agents-hosting';

// Resource identifiers for downstream APIs. These are the RESOURCE only —
// MsalTokenProvider internally appends `/.default` (see msalTokenProvider.ts
// acquireAccessTokenViaSecret: scopes: [`${scope}/.default`]). Passing the
// full `.../.default` scope here would double-suffix and break.
//
// MessagingBotApi uses the URI form per the Bot Connector auth docs (see
// https://learn.microsoft.com/azure/bot-service/rest-api/bot-framework-rest-connector-authentication
// single-tenant section: `scope=https://api.botframework.com/.default`).
// The SDK source also standardizes on this URI at
// Agents-for-js/packages/agents-hosting/src/app/proactive/createConversationOptions.ts
// (`export const AzureBotScope = 'https://api.botframework.com'`).
const MESSAGING_BOT_API_SCOPE = 'https://api.botframework.com';
const GRAPH_SCOPE = 'https://graph.microsoft.com';

// Public Microsoft Entra authority. MsalTokenProvider builds the full token
// endpoint as `${authority}/${tenantId}` (msalTokenProvider.ts:441). Without
// this field set, the authority becomes `undefined/<tenantId>` and MSAL
// throws url_parse_error.
const ENTRA_AUTHORITY = 'https://login.microsoftonline.com';

/** Blueprint credential → AuthConfiguration for MsalTokenProvider. */
export function buildBlueprintAuthConfig(cred: M365Agent365Cred): AuthConfiguration {
	if (cred.inlineCredKind === 'clientSecret') {
		return {
			clientId: cred.blueprintAppId,
			clientSecret: cred.blueprintSecret ?? '',
			tenantId: cred.tenantId,
			authority: ENTRA_AUTHORITY,
		};
	}
	if (cred.inlineCredKind === 'clientCert') {
		// Inline cert transport is not wired up. The SDK's AuthConfiguration
		// uses certPemFile/certKeyFile (filesystem paths), not inline PEM
		// content. Supporting this would require writing the PEM to a temp
		// file on every token acquisition (sec/perf concerns) or switching
		// to @azure/msal-node directly. For cert-based auth, use sidecar
		// transport — the sidecar mounts the cert from a Docker secret and
		// handles MSAL internally.
		throw new Error(
			'Inline transport with clientCert is not supported. Use sidecar transport for certificate-based authentication — the sidecar mounts the PFX from a Docker secret and handles MSAL internally.',
		);
	}
	throw new Error(`Unsupported inlineCredKind: ${cred.inlineCredKind}`);
}

/** Map downstream API name → OAuth2 scope string. */
export function scopeForDownstream(api: string): string {
	switch (api) {
		case 'MessagingBotApi':
			return MESSAGING_BOT_API_SCOPE;
		case 'Graph':
			return GRAPH_SCOPE;
		default:
			throw new Error(`Unknown downstream API: ${api}`);
	}
}
