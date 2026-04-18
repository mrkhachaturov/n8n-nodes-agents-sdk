import type {
	AuthKind,
	M365ClassicBotCred,
	M365Agent365Cred,
	M365Agent365ValidateVia,
} from '../types';
import { verifyJwt } from '../verifyJwt';

function defaultAgent365Issuers(tenantId: string): string[] {
	return [
		`https://login.microsoftonline.com/${tenantId}/v2.0`,
		`https://login.microsoftonline.com/${tenantId}/v2.0/`,
		`https://sts.windows.net/${tenantId}/`,
	];
}

function stripBearer(headerOrToken: string): string {
	return headerOrToken.replace(/^Bearer\s+/i, '');
}

export async function validateInline(
	authKind: AuthKind,
	cred: M365ClassicBotCred | M365Agent365Cred,
	headerOrToken: string,
): Promise<{ claims: Record<string, unknown> }> {
	const raw = stripBearer(headerOrToken);
	if (authKind === 'classicBot') {
		const c = cred as M365ClassicBotCred;
		const claims = await verifyJwt(raw, {
			clientId: c.clientId,
			tenantId: c.tenantId,
		});
		return { claims };
	}
	const c = cred as M365Agent365Cred;
	const extraAudiences = (c.options?.allowedAudiences ?? '')
		.split(',')
		.map((s) => s.trim())
		.filter(Boolean);
	// Per Microsoft Learn (Entra SDK for AgentID — Troubleshooting), the inbound
	// `aud` depends on the Blueprint's `requestedAccessTokenVersion`:
	//   v2            → `{blueprintAppId}` (GUID)
	//   v1 or null    → `api://{blueprintAppId}` (App ID URI)
	// We accept both defaults so a default-configured Blueprint validates
	// without the user needing to fill the Allowed Audiences override.
	const audiences = [c.blueprintAppId, `api://${c.blueprintAppId}`, ...extraAudiences];

	const extraIssuers = (c.options?.allowedIssuers ?? '')
		.split(',')
		.map((s) => s.trim())
		.filter(Boolean);
	const issuers = [...defaultAgent365Issuers(c.tenantId), ...extraIssuers];

	const claims = await verifyJwt(raw, {
		clientId: c.blueprintAppId,
		tenantId: c.tenantId,
		audiences,
		issuers,
	});
	return { claims };
}

export async function validateViaSidecar(
	cred: M365Agent365Cred,
	headerOrToken: string,
): Promise<{ claims: Record<string, unknown> }> {
	const base = (cred.validatorSidecarUrl ?? cred.sidecarUrl ?? '').replace(/\/$/, '');
	if (!base) throw new Error('Sidecar URL is required for validateVia=sidecar');
	const raw = stripBearer(headerOrToken);
	const authHeader = `Bearer ${raw}`;
	const timeoutMs = cred.options?.sidecarTimeoutMs ?? 30000;
	const controller = new AbortController();
	const timer = setTimeout(() => controller.abort(), timeoutMs);
	try {
		const resp = await fetch(`${base}/Validate`, {
			method: 'GET',
			headers: { Authorization: authHeader },
			signal: controller.signal,
		});
		if (!resp.ok) throw new Error(`sidecar /Validate ${resp.status}`);
		const data = (await resp.json()) as { claims: Record<string, unknown> };
		return data;
	} finally {
		clearTimeout(timer);
	}
}

export function effectiveValidateVia(
	authKind: AuthKind,
	cred: M365ClassicBotCred | M365Agent365Cred,
): 'inline' | 'sidecar' {
	if (authKind === 'classicBot') return 'inline';
	const c = cred as M365Agent365Cred;
	const via: M365Agent365ValidateVia = c.validateVia ?? 'sameAsOutbound';
	if (via === 'sameAsOutbound') return c.transport === 'sidecar' ? 'sidecar' : 'inline';
	return via;
}
