import type { M365Agent365Cred } from '../types';
import type { AuthConfiguration } from '@microsoft/agents-hosting';

// Microsoft Bot Framework service app ID. Used as the resource/audience for
// Bot Connector API tokens. Source: sidecar env DownstreamApis__MessagingBotApi__Scopes__0
// in production (Infra/Containers/swarm/stacks/automation/n8n). Spec §7.1, §8.2.
const MESSAGING_BOT_API_SCOPE = 'a6c6ce43-d3ed-40ae-a6d2-4f9c028e0355/.default';
const GRAPH_SCOPE = 'https://graph.microsoft.com/.default';

/** Blueprint credential → AuthConfiguration for MsalTokenProvider. */
export function buildBlueprintAuthConfig(cred: M365Agent365Cred): AuthConfiguration {
	if (cred.inlineCredKind === 'clientSecret') {
		return {
			clientId: cred.blueprintAppId,
			clientSecret: cred.blueprintSecret ?? '',
			tenantId: cred.tenantId,
		};
	}
	if (cred.inlineCredKind === 'clientCert') {
		// Inline cert transport is NOT supported in M1.
		// The SDK's AuthConfiguration uses certPemFile/certKeyFile (filesystem paths),
		// not inline PEM content. Supporting this would require writing the PEM to a
		// temp file on every token acquisition (sec/perf concerns) or switching to
		// @azure/msal-node directly. For cert-based auth, use sidecar transport —
		// the sidecar mounts the cert from a Docker secret and handles MSAL internally.
		throw new Error(
			'Inline transport with clientCert is not supported in M1. Use sidecar transport for certificate-based authentication.',
		);
	}
	throw new Error(`Unsupported inlineCredKind: ${cred.inlineCredKind}`);
}

/** Map downstream API name → OAuth2 scope string. */
export function scopeForDownstream(api: string): string {
	switch (api) {
		case 'MessagingBotApi': return MESSAGING_BOT_API_SCOPE;
		case 'Graph': return GRAPH_SCOPE;
		default: throw new Error(`Unknown downstream API: ${api}`);
	}
}
