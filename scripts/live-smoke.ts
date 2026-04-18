#!/usr/bin/env -S node --import tsx
// Live smoke — skip unless LIVE_SMOKE=1.
//   SIDECAR_URL, BLUEPRINT_APP_ID, TENANT_ID (required)
//   TEST_AGENT_UPN (optional — exercises agentUser path)
//   TEST_DELEGATED_USER_TOKEN (optional — exercises interactiveOBO path)
import { acquireOutboundToken } from '../shared/auth/router';
import type { M365Agent365Cred } from '../shared/types';

if (process.env.LIVE_SMOKE !== '1') {
	console.log('[live-smoke] skipped — set LIVE_SMOKE=1 to run.');
	process.exit(0);
}

const required = (name: string): string => {
	const v = process.env[name];
	if (!v) throw new Error(`${name} required`);
	return v;
};

const cred: M365Agent365Cred = {
	tenantId: required('TENANT_ID'),
	blueprintAppId: required('BLUEPRINT_APP_ID'),
	agentInstanceAppId: process.env.AGENT_INSTANCE_APP_ID, // optional, falls back to blueprintAppId
	transport: 'sidecar',
	sidecarUrl: required('SIDECAR_URL'),
	validateVia: 'sameAsOutbound',
};

async function main() {
	const auto = await acquireOutboundToken({
		authKind: 'agent365',
		credentials: cred,
		identityMode: 'autonomous',
		downstreamApi: 'MessagingBotApi',
	});
	console.log('[live-smoke] autonomous OK, header head:', auto.authorizationHeader.slice(0, 20));

	if (process.env.TEST_AGENT_UPN) {
		const au = await acquireOutboundToken({
			authKind: 'agent365',
			credentials: cred,
			identityMode: 'agentUser',
			downstreamApi: 'MessagingBotApi',
			agentUsername: process.env.TEST_AGENT_UPN,
		});
		console.log('[live-smoke] agent-user OK, header head:', au.authorizationHeader.slice(0, 20));
	}

	if (process.env.TEST_DELEGATED_USER_TOKEN) {
		const obo = await acquireOutboundToken({
			authKind: 'agent365',
			credentials: cred,
			identityMode: 'interactiveOBO',
			downstreamApi: 'Graph',
			inboundBearer: process.env.TEST_DELEGATED_USER_TOKEN,
		});
		console.log('[live-smoke] OBO OK, header head:', obo.authorizationHeader.slice(0, 20));
	}
}

main().catch((e) => {
	console.error('[live-smoke] FAIL', e);
	process.exit(1);
});
