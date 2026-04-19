import { describe, it, expect } from 'vitest';
import { M365AgentTrigger } from '../../../nodes/M365AgentTrigger/M365AgentTrigger.node';
import { KNOWN_INVOKE_NAMES } from '../../../shared/invokeNames';

/**
 * Invoke-name filter tests.
 *
 * The filter is implemented in webhook() and only fires for invoke activities
 * when `activityTypes` includes 'invoke'. The envelope always surfaces
 * `invokeName` when the inbound activity is an invoke with a string `name`.
 *
 * Free-text/unknown-name support is validated by injecting a non-registry value
 * through `getNodeParameter` — the filter logic is list-membership, so anything
 * in the saved config passes regardless of UI-type constraints.
 */
function makeCtx(opts: {
	activityTypes?: string[];
	invokeNames?: string[];
	type: string;
	name?: string;
}) {
	// Minimal-but-valid Activity skeleton — envelope builder requires serviceUrl, channelId, conversation.
	const body: Record<string, unknown> = {
		type: opts.type,
		id: 'a1',
		text: 'x',
		serviceUrl: 'https://smba/',
		channelId: 'msteams',
		from: { id: '29:u1' },
		conversation: { id: '19:c1', conversationType: 'personal' },
		recipient: { id: '28:b1' },
	};
	if (opts.name !== undefined) body.name = opts.name;

	return {
		getRequestObject: () => ({
			headers: {},
			body,
			method: 'POST',
		}),
		getBodyData: () => body,
		getHeaderData: () => ({}),
		getNodeParameter: (name: string) => {
			if (name === 'activityTypes') return opts.activityTypes ?? [];
			if (name === 'invokeNames') return opts.invokeNames ?? [];
			if (name === 'channelFilter') return [];
			if (name === 'authKind') return 'classicBot';
			if (name === 'responseMode') return 'onReceived';
			return undefined;
		},
		getCredentials: async () => ({ clientId: 'c', tenantId: 't', anonymousAllowed: true }),
		helpers: { returnJsonArray: (x: any) => x.map((json: any) => ({ json })) },
		getNode: () => ({ name: 'test' }),
	} as any;
}

describe('M365AgentTrigger invoke-name filter', () => {
	const trigger = new M365AgentTrigger();

	it('emits invokeName on envelope when type=invoke', async () => {
		const res = await trigger.webhook.call(
			makeCtx({ type: 'invoke', name: 'taskModule/fetch' }) as any,
		);
		const envelope = (res.workflowData as any)[0][0].json;
		expect(envelope.invokeName).toBe('taskModule/fetch');
	});

	it('omits invokeName on envelope when type=message', async () => {
		const res = await trigger.webhook.call(makeCtx({ type: 'message' }) as any);
		const envelope = (res.workflowData as any)[0][0].json;
		expect(envelope.invokeName).toBeUndefined();
	});

	it('accepts invoke with matching name when filter configured', async () => {
		const res = await trigger.webhook.call(
			makeCtx({
				activityTypes: ['invoke'],
				invokeNames: ['taskModule/fetch'],
				type: 'invoke',
				name: 'taskModule/fetch',
			}) as any,
		);
		expect((res.workflowData as any)[0].length).toBeGreaterThan(0);
	});

	it('skips invoke with non-matching name when filter configured', async () => {
		const res = await trigger.webhook.call(
			makeCtx({
				activityTypes: ['invoke'],
				invokeNames: ['taskModule/fetch'],
				type: 'invoke',
				name: 'composeExtension/query',
			}) as any,
		);
		expect((res.workflowData as any)[0].length).toBe(0);
	});

	it('allows free-text invoke-name (unknown/future)', async () => {
		const res = await trigger.webhook.call(
			makeCtx({
				activityTypes: ['invoke'],
				invokeNames: ['custom/vendor/name'],
				type: 'invoke',
				name: 'custom/vendor/name',
			}) as any,
		);
		expect((res.workflowData as any)[0].length).toBeGreaterThan(0);
	});

	// Metadata assertion — keeps the Trigger dropdown aligned with the shared registry.
	it('Trigger `invokeNames` property exposes every value in KNOWN_INVOKE_NAMES', () => {
		const instance = new M365AgentTrigger();
		const invokeNamesProp = instance.description.properties.find((p) => p.name === 'invokeNames');
		expect(invokeNamesProp).toBeDefined();
		const uiValues = (invokeNamesProp?.options as any[]).map((o) => o.value).sort();
		expect(uiValues).toEqual([...KNOWN_INVOKE_NAMES].sort());
	});
});
