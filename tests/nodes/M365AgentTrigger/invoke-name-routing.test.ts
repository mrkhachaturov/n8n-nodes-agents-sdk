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
type FilterMode = 'simple' | 'advanced';

function makeCtx(opts: {
	mode?: FilterMode;
	activityTypes?: string[];
	invokeNames?: string[];
	type: string;
	name?: string;
}) {
	const mode: FilterMode = opts.mode ?? 'advanced';
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
			if (name === 'activityFilterMode') return mode;
			if (name === 'activityTypesSimple') return mode === 'simple' ? opts.activityTypes ?? [] : [];
			if (name === 'activityTypesAdvanced')
				return mode === 'advanced' ? opts.activityTypes ?? [] : [];
			if (name === 'invokeNamesSimple') return mode === 'simple' ? opts.invokeNames ?? [] : [];
			if (name === 'invokeNamesAdvanced') return mode === 'advanced' ? opts.invokeNames ?? [] : [];
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

	it('omits invokeName on envelope when type=invoke but name is empty string', async () => {
		const res = await trigger.webhook.call(
			makeCtx({ type: 'invoke', name: '' }) as any,
		);
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

	it('Simple mode: accepts invoke with matching name when filter configured', async () => {
		const res = await trigger.webhook.call(
			makeCtx({
				mode: 'simple',
				activityTypes: ['invoke'],
				invokeNames: ['taskModule/fetch'],
				type: 'invoke',
				name: 'taskModule/fetch',
			}) as any,
		);
		expect((res.workflowData as any)[0].length).toBeGreaterThan(0);
	});

	it('Simple mode: skips invoke with non-matching name when filter configured', async () => {
		const res = await trigger.webhook.call(
			makeCtx({
				mode: 'simple',
				activityTypes: ['invoke'],
				invokeNames: ['taskModule/fetch'],
				type: 'invoke',
				name: 'composeExtension/query',
			}) as any,
		);
		expect((res.workflowData as any)[0].length).toBe(0);
	});

	// Metadata assertions — both Simple and Advanced variants expose the full registry.
	it('Trigger `invokeNamesSimple` and `invokeNamesAdvanced` expose every value in KNOWN_INVOKE_NAMES', () => {
		const instance = new M365AgentTrigger();
		const expected = [...KNOWN_INVOKE_NAMES].sort();

		const simpleProp = instance.description.properties.find((p) => p.name === 'invokeNamesSimple');
		expect(simpleProp).toBeDefined();
		const simpleValues = (simpleProp?.options as any[]).map((o) => o.value).sort();
		expect(simpleValues).toEqual(expected);

		const advancedProp = instance.description.properties.find(
			(p) => p.name === 'invokeNamesAdvanced',
		);
		expect(advancedProp).toBeDefined();
		const advancedValues = (advancedProp?.options as any[]).map((o) => o.value).sort();
		expect(advancedValues).toEqual(expected);
	});

	// Strict invoke gate (D1-D): Advanced invokeNames only visible when
	// activityTypesAdvanced explicitly contains 'invoke'. Empty-list "accept all"
	// must hide the invoke-name picker to keep the UI unambiguous.
	it('strict invoke gate: invokeNamesAdvanced requires activityTypesAdvanced to include "invoke"', () => {
		const instance = new M365AgentTrigger();
		const advancedProp = instance.description.properties.find(
			(p) => p.name === 'invokeNamesAdvanced',
		);
		expect(advancedProp).toBeDefined();
		expect((advancedProp as any).displayOptions).toEqual({
			show: {
				activityFilterMode: ['advanced'],
				activityTypesAdvanced: ['invoke'],
			},
		});
	});

	it('strict invoke gate: invokeNamesSimple requires activityTypesSimple to include "invoke"', () => {
		const instance = new M365AgentTrigger();
		const simpleProp = instance.description.properties.find(
			(p) => p.name === 'invokeNamesSimple',
		);
		expect(simpleProp).toBeDefined();
		expect((simpleProp as any).displayOptions).toEqual({
			show: {
				activityFilterMode: ['simple'],
				activityTypesSimple: ['invoke'],
			},
		});
	});
});
