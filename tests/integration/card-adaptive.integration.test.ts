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

descr('Adaptive Card send — real Azure', () => {
	it('sends a templated Adaptive Card into the configured conversation', async () => {
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
				resource: 'adaptiveCard',
				operation: 'send',
				authKind: 'classicBot',
				conversationSource: 'envelope',
				cardTemplate:
					'{"type":"AdaptiveCard","version":"1.4","body":[{"type":"TextBlock","text":"Integration test — ${title}"}]}',
				bindingData: { title: 'Adaptive live' },
				options: {},
			},
		});

		const result = await node.execute.call(ctx as unknown as IExecuteFunctions);
		const item = result[0][0].json as IDataObject;
		expect(item.cardSendResult).toBeDefined();
		expect(typeof (item.cardSendResult as { id?: string }).id).toBe('string');
	}, 30_000);
});
