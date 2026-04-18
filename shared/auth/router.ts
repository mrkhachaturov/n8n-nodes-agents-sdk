import type {
	AuthKind,
	IdentityMode,
	M365ClassicBotCred,
	M365Agent365Cred,
} from '../types';
import { acquireClassicBotToken } from './backends/classicBot';

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
	throw new Error(`Unsupported authKind: ${args.authKind}`);
}

export async function validateInboundToken(
	_authKind: AuthKind,
	_credentials: M365ClassicBotCred | M365Agent365Cred,
	_bearer: string,
): Promise<{ claims: Record<string, unknown> }> {
	throw new Error('validateInboundToken not yet implemented');
}
