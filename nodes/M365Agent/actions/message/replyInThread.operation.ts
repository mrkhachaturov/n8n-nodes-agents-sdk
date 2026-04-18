import type { IExecuteFunctions, IDataObject, INodeProperties, JsonObject } from 'n8n-workflow';
import { NodeApiError, NodeOperationError } from 'n8n-workflow';
import type { Activity } from '@microsoft/agents-activity';

import { acquireOutboundToken } from '../../../../shared/auth/router';
import {
	createConnectorFromBearer,
	replyInThread as replyInThreadApi,
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
} from '../../../../shared/types';
import { resolveConversationReference } from '../../descriptions/conversationReference';
import { identityModeFields } from './identityModeFields';
import { makeBundleKey } from '../bundleKey';

export const description: INodeProperties[] = [...identityModeFields('replyInThread')];

export async function execute(
	this: IExecuteFunctions,
	itemIndex: number,
	authKind: AuthKind,
	credentials: M365ClassicBotCred | M365Agent365Cred,
	bundles: Map<string, BotConnectorBundle>,
): Promise<IDataObject> {
	const item = this.getInputData()[itemIndex].json as IDataObject;
	const ref = resolveConversationReference(this, itemIndex, item);
	const parentActivityId = this.getNodeParameter('parentActivityId', itemIndex) as string;
	if (!parentActivityId) {
		throw new NodeOperationError(this.getNode(), 'Reply in Thread requires Parent Activity ID.', {
			itemIndex,
		});
	}
	const text = this.getNodeParameter('text', itemIndex) as string;
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

	// ── Reply in thread ───────────────────────────────────────────────────────
	try {
		const result = await replyInThreadApi(bundle, ref.conversation.id, parentActivityId, activity);
		return { ...item, replyInThreadResult: { id: result.id } };
	} catch (err) {
		const e = err as Error;
		throw new NodeApiError(this.getNode(), { message: e.message } as JsonObject, {
			message: `M365 Agent reply-in-thread failed: ${e.message}`,
		});
	}
}
