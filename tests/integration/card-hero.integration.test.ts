import { describe, it, expect } from 'vitest';
import type { IDataObject, IExecuteFunctions } from 'n8n-workflow';
import { M365Agent } from '../../nodes/M365Agent/M365Agent.node';
import { makeExecuteContext } from '../helpers/makeContext';

const clientId = process.env.M365_TEST_CLIENT_ID;
const clientSecret = process.env.M365_TEST_CLIENT_SECRET;
const tenantId = process.env.M365_TEST_TENANT_ID;
const conversationId = process.env.M365_TEST_CONVERSATION_ID;
const serviceUrl = process.env.M365_TEST_SERVICE_URL;

const hasEnv = Boolean(clientId && clientSecret && tenantId && conversationId && serviceUrl);
const descr = hasEnv ? describe : describe.skip;

descr('Hero Card send — real Azure', () => {
	it('sends a Hero Card into the configured conversation and returns an id', async () => {
		const ref = {
			serviceUrl: serviceUrl as string,
			conversation: { id: conversationId as string },
			channelId: 'msteams',
		};
		const node = new M365Agent();
		const ctx = makeExecuteContext({
			inputItems: [{ conversationReference: ref }],
			credentials: {
				appType: 'SingleTenant',
				clientId: clientId as string,
				clientSecret: clientSecret as string,
				tenantId: tenantId as string,
				anonymousAllowed: false,
			},
			parameters: {
				resource: 'heroCard',
				operation: 'send',
				authKind: 'classicBot',
				conversationSource: 'envelope',
				title: 'Integration test',
				subtitle: 'Hero live',
				text: 'Sent from Plan C Task 22',
				images: {},
				buttons: {
					button: [
						{
							type: 'openUrl',
							title: 'Docs',
							value: 'https://aka.ms/agents-sdk',
							options: {},
						},
					],
				},
				options: { fallbackText: 'Hero fallback' },
			},
		});

		const result = await node.execute.call(ctx as unknown as IExecuteFunctions);
		const item = result[0][0].json as IDataObject;
		expect(item.cardSendResult).toBeDefined();
		expect(typeof (item.cardSendResult as { id?: string }).id).toBe('string');
	}, 30_000);
});
