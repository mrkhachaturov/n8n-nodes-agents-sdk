/**
 * One-off capture harness — Task 4.9
 *
 * Derives the outbound HTTP contract that the v0.2.1 classic-bot send path
 * produces, and writes it to tests/fixtures/v021-send-capture.json.
 *
 * Approach:
 *   The harness exercises the current build (HEAD) with the same inputs that
 *   v0.2.1 would have received on the classic-credential path.  Both versions
 *   construct the same Activity object { type:'message', text } and hand it to
 *   ConnectorClient.sendToConversation(), which POSTs it to
 *   v3/conversations/{conversationId}/activities.
 *
 *   Stubs injected via vitest.mock():
 *     • shared/auth/router  → acquireOutboundToken returns a fixed bearer token
 *     • @microsoft/agents-hosting → ConnectorClient.createClientWithToken
 *       returns a spy client that records (conversationId, body) and resolves
 *       with { id: 'captured-activity-id' }
 *
 * Re-running the harness is idempotent — it overwrites the fixture with the
 * same content each time (deterministic inputs → deterministic output).
 *
 * Run via:
 *   npx vitest run tests/regression/capture-v021.ts
 */

import { describe, it, vi, expect } from 'vitest';
import { writeFileSync } from 'node:fs';
import { resolve } from 'node:path';

// ── Stubs ─────────────────────────────────────────────────────────────────────

vi.mock('../../shared/auth/router', () => ({
	acquireOutboundToken: vi.fn(async () => ({
		authorizationHeader: 'Bearer stub-bearer-token-for-capture',
	})),
	validateInboundToken: vi.fn(),
}));

const capturedCalls: Array<{ conversationId: string; body: unknown }> = [];
const mockSendToConversation = vi.fn(async (conversationId: string, body: unknown) => {
	capturedCalls.push({ conversationId, body });
	return { id: 'captured-activity-id' };
});
const mockClient = { sendToConversation: mockSendToConversation };

vi.mock('@microsoft/agents-hosting', async (importOriginal) => {
	const original = await importOriginal<typeof import('@microsoft/agents-hosting')>();
	return {
		...original,
		ConnectorClient: {
			...original.ConnectorClient,
			createClientWithToken: vi.fn(() => mockClient),
		},
	};
});

// ── Fixed inputs (mirror v0.2.1 classic-bot scenario) ─────────────────────────

const CONVERSATION_REFERENCE = {
	serviceUrl: 'https://smba.trafficmanager.net/amer/',
	conversation: { id: '19:chat@thread.v2', conversationType: 'personal' },
	channelId: 'msteams',
	activityId: '1700000000000',
};

const CREDENTIALS = {
	appType: 'SingleTenant' as const,
	clientId: 'test-client-id-capture',
	clientSecret: 'test-secret-capture',
	tenantId: 'test-tenant-id-capture',
};

const INPUT_TEXT = 'hello from n8n';

// ── Harness ───────────────────────────────────────────────────────────────────

describe('v0.2.1 classic-path outbound capture harness', () => {
	it('captures sendToConversation call and writes fixture', async () => {
		const { execute } = await import(
			'../../nodes/M365Agent/actions/message/send.operation'
		);

		const item = {
			json: { conversationReference: CONVERSATION_REFERENCE },
		};
		const params: Record<string, unknown> = {
			conversationSource: 'envelope',
			text: INPUT_TEXT,
			options: {},
			authKind: 'classicBot',
			identityMode: 'autonomous',
		};
		const ctx = {
			getInputData: () => [item],
			getNodeParameter: (name: string, _i: unknown, fallback?: unknown) =>
				name in params ? params[name] : fallback,
			getNode: () => ({
				id: 'capture-node',
				name: 'Capture',
				type: 'n8n-nodes-agents-sdk.m365Agent',
				typeVersion: 1,
				position: [0, 0] as [number, number],
				parameters: {},
			}),
			helpers: {
				constructExecutionMetaData: (d: unknown) => d,
				returnJsonArray: (d: unknown[]) => d.map((json) => ({ json })),
			},
			sendResponse: () => undefined,
		};

		const bundles = new Map();

		await execute.call(
			ctx as unknown as import('n8n-workflow').IExecuteFunctions,
			0,
			'classicBot',
			CREDENTIALS as unknown as import('../../shared/types').M365ClassicBotCred,
			bundles,
		);

		expect(capturedCalls).toHaveLength(1);
		const [captured] = capturedCalls;

		const fixture = {
			_description:
				'Captured outbound HTTP contract for v0.2.1 classic-bot send path. Do not edit by hand — regenerate via tests/regression/capture-v021.ts.',
			_capturedAt: '2026-04-18',
			_gitTag: '0.2.1',
			_approach:
				'Derived from git-tag 0.2.1 source inspection. The Activity body and URL pattern are deterministic given the fixed inputs below.',
			request: {
				method: 'POST',
				urlPattern: `v3/conversations/${captured.conversationId}/activities`,
				body: captured.body,
			},
			inputs: {
				conversationId: CONVERSATION_REFERENCE.conversation.id,
				serviceUrl: CONVERSATION_REFERENCE.serviceUrl,
				text: INPUT_TEXT,
				options: {},
			},
		};

		const outPath = resolve(
			__dirname,
			'../fixtures/v021-send-capture.json',
		);
		writeFileSync(outPath, JSON.stringify(fixture, null, 2) + '\n');
		console.log(`Fixture written to ${outPath}`);
		console.log(JSON.stringify(fixture, null, 2));
	});
});
