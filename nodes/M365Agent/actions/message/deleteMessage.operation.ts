import type { IExecuteFunctions, IDataObject, INodeProperties, JsonObject } from 'n8n-workflow';
import { NodeApiError, NodeOperationError } from 'n8n-workflow';

import { buildAuthConfig } from '../../../../shared/buildAuthConfig';
import { createConnector, type BotConnectorBundle } from '../../../../shared/botConnector';
import type { M365AgentCredentials } from '../../../../shared/types';
import { resolveConversationReference } from '../../descriptions/conversationReference';
import { identityModeFields } from './identityModeFields';

export const description: INodeProperties[] = [
	...identityModeFields('delete'),
];

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
			'Delete requires activityId in the conversation reference.',
			{ itemIndex },
		);
	}

	let bundle = bundles.get(ref.serviceUrl);
	if (!bundle) {
		bundle = await createConnector(buildAuthConfig(creds), ref.serviceUrl);
		bundles.set(ref.serviceUrl, bundle);
	}

	try {
		await bundle.client.deleteActivity(ref.conversation.id, ref.activityId);
		return { ...item, deleteResult: { deleted: true } };
	} catch (err) {
		const e = err as Error;
		throw new NodeApiError(ctx.getNode(), { message: e.message } as JsonObject, {
			message: `M365 Agent delete failed: ${e.message}`,
		});
	}
}
