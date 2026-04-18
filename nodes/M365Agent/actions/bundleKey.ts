import type { AuthKind, IdentityMode } from '../../../shared/types';

/**
 * Derive a map key for the per-execution connector-bundle cache.
 * Kept in its own module so individual operation files can import it
 * without pulling in the full router (which would create a circular dep
 * through the message/card index modules).
 */
export function makeBundleKey(
	serviceUrl: string,
	authKind: AuthKind,
	identityMode: IdentityMode,
	agentUsername?: string,
	agentUserId?: string,
): string {
	return [serviceUrl, authKind, identityMode, agentUsername ?? '', agentUserId ?? ''].join('::');
}
