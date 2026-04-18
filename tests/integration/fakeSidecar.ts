import express, { type Express, type Request } from 'express';
import { AddressInfo } from 'node:net';
import { Server } from 'node:http';

export interface RecordedCall {
	method: string;
	path: string;
	query: Request['query'];
	authorization?: string;
}

export interface FakeSidecar {
	url: string;
	calls: RecordedCall[];
	stop: () => Promise<void>;
	app: Express;
}

/** Mirrors the real Entra auth sidecar surface exercised in Phase 0. */
export async function startFakeSidecar(): Promise<FakeSidecar> {
	const app = express();
	const calls: RecordedCall[] = [];

	app.use((req, _res, next) => {
		calls.push({
			method: req.method,
			path: req.path,
			query: req.query,
			authorization: req.header('authorization') ?? undefined,
		});
		next();
	});

	app.get('/AuthorizationHeaderUnauthenticated/:downstreamApi', (req, res) => {
		if (!req.query.AgentIdentity) return res.status(400).json({ error: 'AgentIdentity required' });
		if (req.params.downstreamApi === 'NotConfigured')
			return res.status(404).json({ error: 'unknown downstream' });
		if (req.params.downstreamApi === 'FailUpstream')
			return res.status(500).json({ error: 'upstream_failure' });
		const suffix = req.query.AgentUsername
			? 'agent-user'
			: req.query.AgentUserId
				? 'agent-user-oid'
				: 'autonomous';
		res.json({
			authorizationHeader: `Bearer fake.${suffix}.${req.params.downstreamApi}.jwt`,
		});
	});

	app.get('/AuthorizationHeader/:downstreamApi', (req, res) => {
		if (!req.header('authorization')?.startsWith('Bearer ')) {
			return res.status(401).json({ error: 'inbound bearer required' });
		}
		res.json({ authorizationHeader: `Bearer fake.obo.${req.params.downstreamApi}.jwt` });
	});

	app.get('/Validate', (req, res) => {
		const bearer = req.header('authorization');
		if (!bearer || bearer === 'Bearer not.a.valid.jwt') {
			return res.status(401).json({ error: 'invalid_token' });
		}
		res.json({ claims: { aud: 'messaging-bot-api', iss: 'https://api.botframework.com' } });
	});

	return await new Promise((resolve) => {
		const server: Server = app.listen(0, '127.0.0.1', () => {
			const { port } = server.address() as AddressInfo;
			resolve({
				url: `http://127.0.0.1:${port}`,
				calls,
				app,
				stop: () => new Promise<void>((r) => server.close(() => r())),
			});
		});
	});
}
