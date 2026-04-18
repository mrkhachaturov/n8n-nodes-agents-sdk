import type { IExecuteFunctions, IDataObject, INodeProperties, JsonObject } from 'n8n-workflow';
import { NodeApiError, NodeOperationError } from 'n8n-workflow';
import type { Activity } from '@microsoft/agents-activity';

import { buildAuthConfig } from '../../../../shared/buildAuthConfig';
import { createConnector, type BotConnectorBundle } from '../../../../shared/botConnector';
import type { M365AgentCredentials } from '../../../../shared/types';
import { resolveConversationReference } from '../../descriptions/conversationReference';
import { identityModeFields } from './identityModeFields';

export const description: INodeProperties[] = [
	...identityModeFields('update'),
];

export async function execute(
	ctx: IExecuteFunctions,
	itemIndex: number,
	creds: M365AgentCredentials,
	bundles: Map<string, BotConnectorBundle>,
): Promise<IDataObject> {
	const item = ctx.getInputData()[itemIndex].json as IDataObject;
	const ref = resolveConversationReference(ctx, itemIndex, item);
	const text = ctx.getNodeParameter('text', itemIndex) as string;
	const options = ctx.getNodeParameter('options', itemIndex, {}) as IDataObject;

	if (!ref.activityId) {
		throw new NodeOperationError(
			ctx.getNode(),
			'Update requires activityId in the conversation reference.',
			{ itemIndex },
		);
	}

	let renderedText = text;
	if (options.workflowFooter) {
		renderedText = `${text}\n\n_Sent from n8n workflow._`;
	}

	const activity: Partial<Activity> = { type: 'message', text: renderedText };

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
		return { ...item, updateResult: { id: result.id } };
	} catch (err) {
		const e = err as Error;
		throw new NodeApiError(ctx.getNode(), { message: e.message } as JsonObject, {
			message: `M365 Agent update failed: ${e.message}`,
		});
	}
}
