import type { M365Agent365Cred, IdentityMode } from '../../types';

interface SidecarResponse {
	authorizationHeader: string;
}

export async function acquireAgent365SidecarToken(
	cred: M365Agent365Cred,
	identityMode: IdentityMode,
	downstreamApi: string,
	agentUsername?: string,
	agentUserId?: string,
	inboundBearer?: string,
): Promise<{ authorizationHeader: string }> {
	if (!cred.sidecarUrl) {
		throw new Error('sidecarUrl is required for sidecar transport');
	}
	const base = cred.sidecarUrl.replace(/\/$/, '');
	const timeoutMs = cred.options?.sidecarTimeoutMs ?? 30000;

	const isOBO = identityMode === 'interactiveOBO';
	const endpoint = isOBO
		? `/AuthorizationHeader/${encodeURIComponent(downstreamApi)}`
		: `/AuthorizationHeaderUnauthenticated/${encodeURIComponent(downstreamApi)}`;

	// AgentIdentity: the Agent Identity instance appId (distinct from the
	// blueprint's appId — the blueprint is the sidecar's OAuth client, the
	// instance is the per-call identity). Falls back to blueprintAppId for
	// back-compat; Task 2.1's credential UI surfaces both fields.
	const agentIdentity = cred.agentInstanceAppId ?? cred.blueprintAppId;
	const params = new URLSearchParams({ AgentIdentity: agentIdentity });
	if (identityMode === 'agentUser') {
		if (agentUsername) params.set('AgentUsername', agentUsername);
		else if (agentUserId) params.set('AgentUserId', agentUserId);
		else throw new Error('agentUser mode requires agentUsername or agentUserId');
	}

	const headers: Record<string, string> = {};
	if (isOBO) {
		if (!inboundBearer) throw new Error('interactiveOBO requires inboundBearer');
		headers.Authorization = inboundBearer;
	}

	const url = `${base}${endpoint}?${params.toString()}`;
	const controller = new AbortController();
	const timer = setTimeout(() => controller.abort(), timeoutMs);

	let response: Response;
	try {
		response = await fetch(url, { method: 'GET', headers, signal: controller.signal });
	} finally {
		clearTimeout(timer);
	}

	if (!response.ok) {
		const body = await response.text().catch(() => '');
		throw new Error(`sidecar ${response.status} at ${url}: ${body.slice(0, 500)}`);
	}
	const data = (await response.json()) as SidecarResponse;
	if (!data.authorizationHeader) {
		throw new Error(`sidecar response missing authorizationHeader: ${JSON.stringify(data)}`);
	}
	return { authorizationHeader: data.authorizationHeader };
}
