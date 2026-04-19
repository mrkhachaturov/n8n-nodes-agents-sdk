import { describe, it, expect, vi, beforeEach } from 'vitest';
import * as cardAdaptive from '../../../nodes/M365Agent/actions/card-adaptive';
import { M365Agent } from '../../../nodes/M365Agent/M365Agent.node';
import { makeExecuteContext, makeCredentials } from '../../helpers/makeContext';
import type { IExecuteFunctions } from 'n8n-workflow';

// ---------------------------------------------------------------------------
// Mock the auth router and bot connector — no real network calls
// ---------------------------------------------------------------------------

vi.mock('../../../shared/auth/router', () => ({
	acquireOutboundToken: vi.fn().mockResolvedValue({ authorizationHeader: 'Bearer fake-token' }),
}));

const fakeClient = {
	replyToActivity: vi.fn(),
	sendToConversation: vi.fn(),
	updateActivity: vi.fn(),
	deleteActivity: vi.fn(),
};

const fakeBundle = {
	client: fakeClient,
	axios: {},
	token: 'fake-token',
	baseURL: 'https://smba.trafficmanager.net/emea/',
};

const mockCreateConnectorFromBearer = vi.fn().mockReturnValue(fakeBundle);

vi.mock('../../../shared/botConnector', () => ({
	createConnectorFromBearer: (...args: unknown[]) => mockCreateConnectorFromBearer(...args),
	replyInThread: vi.fn(),
}));

// ---------------------------------------------------------------------------
// Description-level assertions
// ---------------------------------------------------------------------------

describe('adaptiveCard resource description', () => {
	it('exposes Send / Update in alphabetical display order', () => {
		const op = cardAdaptive.description.find((p) => p.name === 'operation');
		const names = (op?.options as { name: string }[] | undefined)?.map((o) => o.name);
		expect(names).toEqual(['Send', 'Update']);
	});

	it('exposes both adaptiveCard operations (alphabetical by value)', () => {
		const op = cardAdaptive.description.find((p) => p.name === 'operation');
		const values = (op?.options as { value: string }[] | undefined)?.map((o) => o.value);
		expect(values).toEqual(['send', 'update']);
	});

	it('each operation has an action string for the subtitle', () => {
		const op = cardAdaptive.description.find((p) => p.name === 'operation');
		const actions = (op?.options as { action?: string }[] | undefined)?.map((o) => o.action);
		expect(actions?.every((a) => typeof a === 'string' && a.length > 0)).toBe(true);
	});

	it('operation selector is noDataExpression and scoped to adaptiveCard resource', () => {
		const op = cardAdaptive.description.find((p) => p.name === 'operation');
		expect(op?.noDataExpression).toBe(true);
		const shown = (op?.displayOptions?.show as Record<string, string[]> | undefined)?.resource;
		expect(shown).toEqual(['adaptiveCard']);
		expect(op?.default).toBe('send');
	});

	it('cardTemplate field shown for both adaptiveCard/send and adaptiveCard/update', () => {
		const tmpl = cardAdaptive.description.find((p) => p.name === 'cardTemplate');
		const shown = tmpl?.displayOptions?.show as Record<string, string[]> | undefined;
		expect(shown?.resource).toEqual(['adaptiveCard']);
		expect(shown?.operation).toEqual(['send', 'update']);
	});

	it('bindingData and options are shown for both send and update (parity with the template)', () => {
		for (const name of ['bindingData', 'options']) {
			const prop = cardAdaptive.description.find((p) => p.name === name);
			const shown = prop?.displayOptions?.show as Record<string, string[]> | undefined;
			expect(shown?.operation).toEqual(['send', 'update']);
		}
	});
});

// ---------------------------------------------------------------------------
// Execute-level: card/update dispatches to client.updateActivity with the
// rendered Adaptive Card attached. Mirrors message/update's contract.
// ---------------------------------------------------------------------------

const CARD_UPDATE_ENVELOPE = {
	conversationReference: {
		serviceUrl: 'https://smba.trafficmanager.net/emea/',
		conversation: { id: '19:sample@thread.tacv2' },
		activityId: 'act-card-to-update-001',
		channelId: 'msteams',
	},
	activity: { type: 'message' },
	// Ancillary field — must survive round-trip per manifest §12
	customWorkflowField: 'preserved',
};

describe('M365Agent adaptiveCard/update execute', () => {
	beforeEach(() => {
		vi.clearAllMocks();
		mockCreateConnectorFromBearer.mockReturnValue(fakeBundle);
		fakeClient.updateActivity.mockResolvedValue({ id: 'updated-activity-id' });
	});

	it('calls client.updateActivity with (conversationId, activityId, activity) and expands the card template', async () => {
		const node = new M365Agent();
		const ctx = makeExecuteContext({
			inputItems: [{ ...CARD_UPDATE_ENVELOPE, title: 'Job #42' }],
			credentials: makeCredentials(),
			parameters: {
				resource: 'adaptiveCard',
				operation: 'update',
				conversationSource: 'envelope',
				cardTemplate:
					'{"type":"AdaptiveCard","version":"1.4","body":[{"type":"TextBlock","text":"${title}"}]}',
				bindingData: { title: 'Job #42' },
				options: { fallbackText: 'Job #42 updated' },
			},
		});

		const result = await node.execute.call(ctx as unknown as IExecuteFunctions);

		expect(fakeClient.updateActivity).toHaveBeenCalledTimes(1);
		const [convId, actId, activity] = fakeClient.updateActivity.mock.calls[0];
		expect(convId).toBe(CARD_UPDATE_ENVELOPE.conversationReference.conversation.id);
		expect(actId).toBe(CARD_UPDATE_ENVELOPE.conversationReference.activityId);
		expect(activity.type).toBe('message');
		expect(activity.text).toBe('Job #42 updated');
		expect(activity.attachments).toHaveLength(1);
		expect(activity.attachments[0].contentType).toBe('application/vnd.microsoft.card.adaptive');
		// Template expansion rendered ${title} from bindingData
		expect(JSON.stringify(activity.attachments[0].content)).toContain('Job #42');

		// Output preserves envelope + ancillary fields, adds cardUpdateResult
		expect(result).toHaveLength(1);
		expect(result[0]).toHaveLength(1);
		const outItem = result[0][0];
		expect(outItem.pairedItem).toEqual({ item: 0 });
		expect(outItem.json).toMatchObject({
			conversationReference: CARD_UPDATE_ENVELOPE.conversationReference,
			customWorkflowField: 'preserved',
			cardUpdateResult: { id: 'updated-activity-id' },
		});
	});

	it('throws NodeOperationError when conversationReference has no activityId', async () => {
		const node = new M365Agent();
		const envelopeNoActId = {
			conversationReference: {
				serviceUrl: 'https://smba.trafficmanager.net/emea/',
				conversation: { id: '19:sample@thread.tacv2' },
				channelId: 'msteams',
				// activityId missing
			},
			activity: { type: 'message' },
		};
		const ctx = makeExecuteContext({
			inputItems: [envelopeNoActId],
			credentials: makeCredentials(),
			parameters: {
				resource: 'adaptiveCard',
				operation: 'update',
				conversationSource: 'envelope',
				cardTemplate: '{"type":"AdaptiveCard","version":"1.4","body":[]}',
				bindingData: {},
				options: {},
			},
		});

		await expect(node.execute.call(ctx as unknown as IExecuteFunctions)).rejects.toThrow(
			/activityId/,
		);
		expect(fakeClient.updateActivity).not.toHaveBeenCalled();
	});
});
