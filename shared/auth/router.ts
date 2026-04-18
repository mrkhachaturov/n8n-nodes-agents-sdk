import type {
	AuthKind,
	IdentityMode,
	M365ClassicBotCred,
	M365Agent365Cred,
} from '../types';
import { acquireClassicBotToken } from './backends/classicBot';
import { acquireAgent365InlineToken } from './backends/agent365Inline';
import { acquireAgent365SidecarToken } from './backends/agent365Sidecar';
import { effectiveValidateVia, validateInline, validateViaSidecar } from './validate';

export interface AcquireOutboundTokenArgs {
	authKind: AuthKind;
	credentials: M365ClassicBotCred | M365Agent365Cred;
	identityMode: IdentityMode;
	downstreamApi: string;
	agentUsername?: string;
	agentUserId?: string;
	inboundBearer?: string;
}

export interface TokenResult {
	authorizationHeader: string;
	expiresAt?: Date;
}

export async function acquireOutboundToken(
	args: AcquireOutboundTokenArgs,
): Promise<TokenResult> {
	if (args.authKind === 'classicBot') {
		return acquireClassicBotToken(args.credentials as M365ClassicBotCred);
	}
	if (args.authKind === 'agent365') {
		const cred = args.credentials as M365Agent365Cred;
		if (cred.transport === 'inline') {
			return acquireAgent365InlineToken(cred, args.identityMode, args.downstreamApi);
		}
		if (cred.transport === 'sidecar') {
			return acquireAgent365SidecarToken(
				cred,
				args.identityMode,
				args.downstreamApi,
				args.agentUsername,
				args.agentUserId,
				args.inboundBearer,
			);
		}
		throw new Error(`Unsupported transport: ${cred.transport}`);
	}
	throw new Error(`Unsupported authKind: ${args.authKind}`);
}

export async function validateInboundToken(
	authKind: AuthKind,
	credentials: M365ClassicBotCred | M365Agent365Cred,
	bearer: string,
): Promise<{ claims: Record<string, unknown> }> {
	const via = effectiveValidateVia(authKind, credentials);
	if (via === 'sidecar') {
		return validateViaSidecar(credentials as M365Agent365Cred, bearer);
	}
	return validateInline(authKind, credentials, bearer);
}
