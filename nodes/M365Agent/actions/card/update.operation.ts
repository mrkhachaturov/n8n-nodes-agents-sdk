import type { IExecuteFunctions, IDataObject, INodeProperties, JsonObject } from 'n8n-workflow';
import { NodeApiError, NodeOperationError } from 'n8n-workflow';
import type { Activity, Attachment } from '@microsoft/agents-activity';
import { Template } from 'adaptivecards-templating';

import { acquireOutboundToken } from '../../../../shared/auth/router';
import { createConnectorFromBearer, type BotConnectorBundle } from '../../../../shared/botConnector';
import type { AuthKind, IdentityMode, M365ClassicBotCred, M365Agent365Cred } from '../../../../shared/types';
import { resolveConversationReference } from '../../descriptions/conversationReference';
import { identityModeFields } from '../identityModeFields';
import { makeBundleKey } from '../bundleKey';

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
	this: IExecuteFunctions,
	itemIndex: number,
	authKind: AuthKind,
	credentials: M365ClassicBotCred | M365Agent365Cred,
	bundles: Map<string, BotConnectorBundle>,
): Promise<IDataObject> {
	const item = this.getInputData()[itemIndex].json as IDataObject;
	const ref = resolveConversationReference(this, itemIndex, item);

	if (!ref.activityId) {
		throw new NodeOperationError(
			this.getNode(),
			'Update requires activityId in the conversation reference.',
			{ itemIndex },
		);
	}

	const cardTemplate = parseJsonParam(
		this,
		itemIndex,
		'Card Template',
		this.getNodeParameter('cardTemplate', itemIndex),
	);
	const bindingData = parseJsonParam(
		this,
		itemIndex,
		'Binding Data',
		this.getNodeParameter('bindingData', itemIndex, {}),
	);

	const options = this.getNodeParameter('options', itemIndex, {}) as IDataObject;

	let rendered: unknown;
	try {
		const tpl = new Template(cardTemplate);
		rendered = tpl.expand({ $root: bindingData });
	} catch (err) {
		throw new NodeOperationError(
			this.getNode(),
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

	// ── Update ────────────────────────────────────────────────────────────────
	try {
		const result = await bundle.client.updateActivity(
			ref.conversation.id,
			ref.activityId,
			activity as Activity,
		);
		return { ...item, cardUpdateResult: { id: result.id } };
	} catch (err) {
		const e = err as Error;
		throw new NodeApiError(this.getNode(), { message: e.message } as JsonObject, {
			message: `M365 Agent card update failed: ${e.message}`,
		});
	}
}
