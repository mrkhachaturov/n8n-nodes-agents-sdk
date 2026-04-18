import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import { startFakeSidecar, type FakeSidecar } from './fakeSidecar';
import { acquireOutboundToken } from '../../shared/auth/router';
import type { M365Agent365Cred } from '../../shared/types';

describe('acquireOutboundToken ↔ fake sidecar (spec §9.3)', () => {
	let sidecar: FakeSidecar;
	let baseCred: M365Agent365Cred;

	beforeAll(async () => {
		sidecar = await startFakeSidecar();
		baseCred = {
			tenantId: 'tid',
			blueprintAppId: 'blueprint-app-id',
			transport: 'sidecar',
			sidecarUrl: sidecar.url,
			validateVia: 'sameAsOutbound',
		} as M365Agent365Cred;
	});
	afterAll(async () => {
		await sidecar.stop();
	});

	it('autonomous calls /AuthorizationHeaderUnauthenticated/{api}?AgentIdentity=…', async () => {
		const result = await acquireOutboundToken({
			authKind: 'agent365',
			credentials: baseCred,
			identityMode: 'autonomous',
			downstreamApi: 'MessagingBotApi',
		});
		expect(result.authorizationHeader).toBe('Bearer fake.autonomous.MessagingBotApi.jwt');

		const call = sidecar.calls.at(-1)!;
		expect(call.path).toBe('/AuthorizationHeaderUnauthenticated/MessagingBotApi');
		expect(call.query.AgentIdentity).toBe('blueprint-app-id');
		expect(call.query.AgentUsername).toBeUndefined();
	});

	it('agentUser adds AgentUsername query param', async () => {
		const result = await acquireOutboundToken({
			authKind: 'agent365',
			credentials: baseCred,
			identityMode: 'agentUser',
			downstreamApi: 'MessagingBotApi',
			agentUsername: 'agent@contoso.com',
		});
		expect(result.authorizationHeader).toBe('Bearer fake.agent-user.MessagingBotApi.jwt');

		const call = sidecar.calls.at(-1)!;
		expect(call.query.AgentUsername).toBe('agent@contoso.com');
	});

	it('interactiveOBO uses /AuthorizationHeader (no Unauthenticated) and forwards inbound bearer', async () => {
		// inboundBearer must include the "Bearer " scheme prefix — the backend
		// forwards it verbatim as the Authorization header, and the fake sidecar
		// requires startsWith('Bearer ') to accept the call.
		const result = await acquireOutboundToken({
			authKind: 'agent365',
			credentials: baseCred,
			identityMode: 'interactiveOBO',
			downstreamApi: 'Graph',
			inboundBearer: 'Bearer inbound.user.jwt',
		});
		expect(result.authorizationHeader).toBe('Bearer fake.obo.Graph.jwt');

		const call = sidecar.calls.at(-1)!;
		expect(call.path).toBe('/AuthorizationHeader/Graph');
		expect(call.authorization).toBe('Bearer inbound.user.jwt');
	});

	it('propagates sidecar 500 as error', async () => {
		await expect(
			acquireOutboundToken({
				authKind: 'agent365',
				credentials: baseCred,
				identityMode: 'autonomous',
				downstreamApi: 'FailUpstream',
			}),
		).rejects.toThrow(/sidecar.*500/i);
	});
});
