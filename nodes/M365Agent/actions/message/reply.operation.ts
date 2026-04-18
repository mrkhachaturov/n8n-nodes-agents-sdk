import type { IExecuteFunctions, IDataObject, INodeProperties, JsonObject } from 'n8n-workflow';
import { NodeApiError, NodeOperationError } from 'n8n-workflow';
import type { Activity } from '@microsoft/agents-activity';

import { acquireOutboundToken } from '../../../../shared/auth/router';
import { createConnectorFromBearer, type BotConnectorBundle } from '../../../../shared/botConnector';
import type { AuthKind, IdentityMode, M365ClassicBotCred, M365Agent365Cred } from '../../../../shared/types';
import { resolveConversationReference } from '../../descriptions/conversationReference';
import { identityModeFields } from './identityModeFields';
import { makeBundleKey } from '../bundleKey';

export const description: INodeProperties[] = [
	...identityModeFields('reply'),
];

export async function execute(
	this: IExecuteFunctions,
	itemIndex: number,
	authKind: AuthKind,
	credentials: M365ClassicBotCred | M365Agent365Cred,
	bundles: Map<string, BotConnectorBundle>,
): Promise<IDataObject> {
	const item = this.getInputData()[itemIndex].json as IDataObject;
	const ref = resolveConversationReference(this, itemIndex, item);
	const text = this.getNodeParameter('text', itemIndex) as string;
	const options = this.getNodeParameter('options', itemIndex, {}) as IDataObject;

	if (!ref.activityId) {
		throw new NodeOperationError(
			this.getNode(),
			'Reply requires activityId in the conversation reference. The inbound trigger envelope includes it automatically; for manual source, fill Activity ID.',
			{ itemIndex },
		);
	}

	let renderedText = text;
	if (options.workflowFooter) {
		renderedText = `${text}\n\n_Sent from n8n workflow._`;
	}

	const activity: Partial<Activity> = { type: 'message', text: renderedText };

	// ── Auth routing ──────────────────────────────────────────────────────────
	const identityMode: IdentityMode = authKind === 'agent365'
		? (this.getNodeParameter('identityMode', itemIndex, 'autonomous') as IdentityMode)
		: 'autonomous';

	let agentUsername: string | undefined;
	let agentUserId: string | undefined;
	if (authKind === 'agent365' && identityMode === 'agentUser') {
		const userSelectorMode = this.getNodeParameter('userSelectorMode', itemIndex, 'byUpn') as string;
		if (userSelectorMode === 'byUpn') {
			agentUsername = this.getNodeParameter('agentUsername', itemIndex, '') as string || undefined;
		} else {
			agentUserId = this.getNodeParameter('agentUserId', itemIndex, '') as string || undefined;
		}
	}

	const downstreamApi: string = authKind === 'agent365'
		? ((credentials as M365Agent365Cred).outboundDownstreamApi ?? 'MessagingBotApi')
		: 'BotFramework';

	// ── Bundle cache ──────────────────────────────────────────────────────────
	const bundleKey = makeBundleKey(ref.serviceUrl, authKind, identityMode, agentUsername, agentUserId);
	let bundle = bundles.get(bundleKey);
	if (!bundle) {
		const { authorizationHeader } = await acquireOutboundToken({
			authKind,
			credentials,
			identityMode,
			downstreamApi,
			agentUsername,
			agentUserId,
		});
		bundle = createConnectorFromBearer(ref.serviceUrl, authorizationHeader);
		bundles.set(bundleKey, bundle);
	}

	// ── Reply ─────────────────────────────────────────────────────────────────
	try {
		const result = await bundle.client.replyToActivity(
			ref.conversation.id,
			ref.activityId,
			activity as Activity,
		);
		return { ...item, replyResult: { id: result.id } };
	} catch (err) {
		const e = err as Error;
		throw new NodeApiError(this.getNode(), { message: e.message } as JsonObject, {
			message: `M365 Agent reply failed: ${e.message}`,
		});
	}
}
