import type { IExecuteFunctions, IDataObject, INodeProperties, JsonObject } from 'n8n-workflow';
import { NodeApiError } from 'n8n-workflow';
import type { Activity, Attachment } from '@microsoft/agents-activity';
import { Template } from 'adaptivecards-templating';

import { buildAuthConfig } from '../../../../shared/buildAuthConfig';
import { createConnector, type BotConnectorBundle } from '../../../../shared/botConnector';
import type { M365AgentCredentials } from '../../../../shared/types';
import { resolveConversationReference } from '../../descriptions/conversationReference';

export const description: INodeProperties[] = [];

export async function execute(
	ctx: IExecuteFunctions,
	itemIndex: number,
	creds: M365AgentCredentials,
	bundles: Map<string, BotConnectorBundle>,
): Promise<IDataObject> {
	const item = ctx.getInputData()[itemIndex].json as IDataObject;
	const ref = resolveConversationReference(ctx, itemIndex, item);

	// cardTemplate may come in as a string or as an already-parsed object depending
	// on how n8n serialized the json parameter. Normalize to an object.
	const cardTemplateRaw = ctx.getNodeParameter('cardTemplate', itemIndex) as unknown;
	const cardTemplate: Record<string, unknown> =
		typeof cardTemplateRaw === 'string'
			? (JSON.parse(cardTemplateRaw) as Record<string, unknown>)
			: (cardTemplateRaw as Record<string, unknown>);

	const bindingDataRaw = ctx.getNodeParameter('bindingData', itemIndex, {}) as unknown;
	const bindingData: Record<string, unknown> =
		typeof bindingDataRaw === 'string'
			? (JSON.parse(bindingDataRaw) as Record<string, unknown>)
			: (bindingDataRaw as Record<string, unknown>);

	const options = ctx.getNodeParameter('options', itemIndex, {}) as IDataObject;

	const tpl = new Template(cardTemplate);
	const rendered = tpl.expand({ $root: bindingData });

	const attachment: Attachment = {
		contentType: 'application/vnd.microsoft.card.adaptive',
		content: rendered,
	};

	const activity: Partial<Activity> = {
		type: 'message',
		attachments: [attachment],
		...(options.fallbackText ? { text: options.fallbackText as string } : {}),
	};

	let bundle = bundles.get(ref.serviceUrl);
	if (!bundle) {
		bundle = await createConnector(buildAuthConfig(creds), ref.serviceUrl);
		bundles.set(ref.serviceUrl, bundle);
	}

	try {
		const result = await bundle.client.sendToConversation(ref.conversation.id, activity as Activity);
		return { ...item, cardSendResult: { id: result.id } };
	} catch (err) {
		const e = err as Error;
		throw new NodeApiError(ctx.getNode(), { message: e.message } as JsonObject, {
			message: `M365 Agent card send failed: ${e.message}`,
		});
	}
}
