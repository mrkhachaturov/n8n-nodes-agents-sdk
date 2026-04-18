import type { IExecuteFunctions, IDataObject, INodeProperties, JsonObject } from 'n8n-workflow';
import { NodeApiError, NodeOperationError } from 'n8n-workflow';
import type { Activity, Attachment } from '@microsoft/agents-activity';
import { Template } from 'adaptivecards-templating';

import { buildAuthConfig } from '../../../../shared/buildAuthConfig';
import { createConnector, type BotConnectorBundle } from '../../../../shared/botConnector';
import type { M365AgentCredentials } from '../../../../shared/types';
import { resolveConversationReference } from '../../descriptions/conversationReference';
import { identityModeFields } from '../identityModeFields';

export const description: INodeProperties[] = [
	...identityModeFields('card', 'update'),
];

function parseJsonParam(
	ctx: IExecuteFunctions,
	itemIndex: number,
	fieldName: string,
	raw: unknown,
): Record<string, unknown> {
	if (typeof raw !== 'string') {
		return raw as Record<string, unknown>;
	}
	try {
		return JSON.parse(raw) as Record<string, unknown>;
	} catch (err) {
		throw new NodeOperationError(
			ctx.getNode(),
			`Invalid JSON in ${fieldName}: ${(err as Error).message}`,
			{ itemIndex },
		);
	}
}

export async function execute(
	ctx: IExecuteFunctions,
	itemIndex: number,
	creds: M365AgentCredentials,
	bundles: Map<string, BotConnectorBundle>,
): Promise<IDataObject> {
	const item = ctx.getInputData()[itemIndex].json as IDataObject;
	const ref = resolveConversationReference(ctx, itemIndex, item);

	if (!ref.activityId) {
		throw new NodeOperationError(
			ctx.getNode(),
			'Update requires activityId in the conversation reference.',
			{ itemIndex },
		);
	}

	const cardTemplate = parseJsonParam(
		ctx,
		itemIndex,
		'Card Template',
		ctx.getNodeParameter('cardTemplate', itemIndex),
	);
	const bindingData = parseJsonParam(
		ctx,
		itemIndex,
		'Binding Data',
		ctx.getNodeParameter('bindingData', itemIndex, {}),
	);

	const options = ctx.getNodeParameter('options', itemIndex, {}) as IDataObject;

	let rendered: unknown;
	try {
		const tpl = new Template(cardTemplate);
		rendered = tpl.expand({ $root: bindingData });
	} catch (err) {
		throw new NodeOperationError(
			ctx.getNode(),
			`Card Template expansion failed: ${(err as Error).message}`,
			{ itemIndex },
		);
	}

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
		const result = await bundle.client.updateActivity(
			ref.conversation.id,
			ref.activityId,
			activity as Activity,
		);
		return { ...item, cardUpdateResult: { id: result.id } };
	} catch (err) {
		const e = err as Error;
		throw new NodeApiError(ctx.getNode(), { message: e.message } as JsonObject, {
			message: `M365 Agent card update failed: ${e.message}`,
		});
	}
}
