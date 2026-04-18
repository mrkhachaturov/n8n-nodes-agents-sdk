import type { IExecuteFunctions, IDataObject, INodeProperties, JsonObject } from 'n8n-workflow';
import { NodeApiError, NodeOperationError } from 'n8n-workflow';
import type { Activity } from '@microsoft/agents-activity';

import { acquireOutboundToken } from '../../../../shared/auth/router';
import {
	createConnectorFromBearer,
	type BotConnectorBundle,
} from '../../../../shared/botConnector';
import { applyMentionsToActivity, type MentionInput } from '../../../../shared/mentions';
import {
	applySuggestedActionsToActivity,
	type SuggestedActionInput,
} from '../../../../shared/suggestedActions';
import type {
	AuthKind,
	IdentityMode,
	M365ClassicBotCred,
	M365Agent365Cred,
	ConversationReference,
} from '../../../../shared/types';
import { identityModeFields } from './identityModeFields';
import { makeBundleKey } from '../bundleKey';

export const description: INodeProperties[] = [
	// All send-specific fields are already in message/index.ts (Text, Options).
	// No extra properties unique to send — left as [] so future send-only
	// fields have a home.
	...identityModeFields('send'),
];

export async function execute(
	this: IExecuteFunctions,
	itemIndex: number,
	authKind: AuthKind,
	credentials: M365ClassicBotCred | M365Agent365Cred,
	bundles: Map<string, BotConnectorBundle>,
): Promise<IDataObject> {
	const item = this.getInputData()[itemIndex].json as IDataObject;

	// ── Conversation reference ─────────────────────────────────────────────────
	// When conversationSource is 'manual', read from UI fields; any other value
	// (including 'envelope' and legacy 'fromEnvelope') reads from the item.
	const conversationSource = this.getNodeParameter(
		'conversationSource',
		itemIndex,
		'envelope',
	) as string;
	let ref: ConversationReference;
	if (conversationSource === 'manual') {
		const serviceUrl = this.getNodeParameter('serviceUrl', itemIndex) as string;
		const conversationId = this.getNodeParameter('conversationId', itemIndex) as string;
		const channelId = this.getNodeParameter('channelId', itemIndex, 'msteams') as string;
		const activityId = this.getNodeParameter('activityId', itemIndex, '') as string;
		if (!serviceUrl || !conversationId) {
			throw new NodeOperationError(
				this.getNode(),
				'Manual Conversation Source requires both Service URL and Conversation ID.',
				{ itemIndex },
			);
		}
		ref = {
			serviceUrl,
			conversation: { id: conversationId },
			channelId,
			...(activityId ? { activityId } : {}),
		};
	} else {
		const fromItem = item.conversationReference as ConversationReference | undefined;
		if (!fromItem) {
			throw new NodeOperationError(
				this.getNode(),
				'conversationReference is missing on the input item. Switch Conversation Source to "Specify Manually" or route through M365 Agent Trigger.',
				{ itemIndex },
			);
		}
		ref = fromItem;
	}

	// ── Message text ──────────────────────────────────────────────────────────
	const text = this.getNodeParameter('text', itemIndex, '') as string;
	const options = this.getNodeParameter('options', itemIndex, {}) as IDataObject;

	let renderedText = text;
	if (options.workflowFooter) {
		renderedText = `${text}\n\n_Sent from n8n workflow._`;
	}

	let activity: Partial<Activity> = { type: 'message', text: renderedText };

	const mentionsCollection = (options.mentions as { values?: MentionInput[] } | undefined)?.values;
	if (mentionsCollection && mentionsCollection.length > 0) {
		try {
			activity = applyMentionsToActivity(activity, mentionsCollection);
		} catch (err) {
			throw new NodeOperationError(this.getNode(), (err as Error).message, { itemIndex });
		}
	}

	const actionsCollection = (
		options.suggestedActions as { values?: SuggestedActionInput[] } | undefined
	)?.values;
	if (actionsCollection && actionsCollection.length > 0) {
		try {
			activity = applySuggestedActionsToActivity(activity, actionsCollection);
		} catch (err) {
			throw new NodeOperationError(this.getNode(), (err as Error).message, { itemIndex });
		}
	}

	// ── Auth routing ──────────────────────────────────────────────────────────
	// Determine identity context from node params (shown only when authKind=agent365).
	const identityMode: IdentityMode =
		authKind === 'agent365'
			? (this.getNodeParameter('identityMode', itemIndex, 'autonomous') as IdentityMode)
			: 'autonomous';

	let agentUsername: string | undefined;
	let agentUserId: string | undefined;
	if (authKind === 'agent365' && identityMode === 'agentUser') {
		const userSelectorMode = this.getNodeParameter(
			'userSelectorMode',
			itemIndex,
			'byUpn',
		) as string;
		if (userSelectorMode === 'byUpn') {
			agentUsername =
				(this.getNodeParameter('agentUsername', itemIndex, '') as string) || undefined;
		} else {
			agentUserId = (this.getNodeParameter('agentUserId', itemIndex, '') as string) || undefined;
		}
	}

	const downstreamApi: string =
		authKind === 'agent365'
			? ((credentials as M365Agent365Cred).outboundDownstreamApi ?? 'MessagingBotApi')
			: 'BotFramework';

	// ── Bundle cache ──────────────────────────────────────────────────────────
	const bundleKey = makeBundleKey(
		ref.serviceUrl,
		authKind,
		identityMode,
		agentUsername,
		agentUserId,
	);
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

	// ── Send ──────────────────────────────────────────────────────────────────
	try {
		const result = await bundle.client.sendToConversation(
			ref.conversation.id,
			activity as Activity,
		);
		return { ...item, sendResult: { id: result.id } };
	} catch (err) {
		const e = err as Error;
		throw new NodeApiError(this.getNode(), { message: e.message } as JsonObject, {
			message: `M365 Agent send failed: ${e.message}`,
		});
	}
}
