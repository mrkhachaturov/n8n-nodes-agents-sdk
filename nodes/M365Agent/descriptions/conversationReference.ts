import type { IExecuteFunctions, IDataObject, INodeProperties } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import type { ConversationReference } from '../../../shared/types';

/**
 * Simple vs Advanced conversation-target input.
 * Simple (default): reads the envelope's conversationReference from the item —
 *   exactly the M0 behaviour, no UI fields shown.
 * Advanced: the user fills serviceUrl / conversation id / channel id / activity id
 *   manually. Used for proactive flows originating outside the bot webhook (e.g.,
 *   a 1C ERP event) where there is no inbound envelope to read.
 *
 * Manifest §6.3 Simple vs Advanced; §7 fallback when the API can't list.
 */
export const conversationReferenceProperties: INodeProperties[] = [
	{
		displayName: 'Conversation Source',
		name: 'conversationSource',
		type: 'options',
		noDataExpression: true,
		options: [
			{
				name: 'From Envelope',
				value: 'envelope',
				description:
					'Use conversationReference from the current item. Works for replies / updates / deletes.',
			},
			{
				name: 'Specify Manually',
				value: 'manual',
				description:
					'Fill the conversation reference fields yourself. Use for proactive sends originating outside the bot.',
			},
		],
		default: 'envelope',
		description:
			'Where the conversation routing info comes from. Reply / update / delete flows use From Envelope automatically; proactive sends need Specify Manually.',
	},
	{
		displayName: 'Service URL',
		name: 'serviceUrl',
		type: 'string',
		default: '',
		required: true,
		displayOptions: { show: { conversationSource: ['manual'] } },
		placeholder: 'https://smba.trafficmanager.net/emea/',
		description: 'Bot Connector service URL for the target channel',
	},
	{
		displayName: 'Conversation ID',
		name: 'conversationId',
		type: 'string',
		default: '',
		required: true,
		displayOptions: { show: { conversationSource: ['manual'] } },
		placeholder: '19:xxx@thread.tacv2',
		description:
			'The conversation ID. For Teams channels this looks like 19:…@thread.tacv2. Paste an expression such as {{ $JSON.channelId }} if you stored it upstream.',
	},
	{
		displayName: 'Channel ID',
		name: 'channelId',
		type: 'options',
		default: 'msteams',
		displayOptions: { show: { conversationSource: ['manual'] } },
		options: [
			{ name: 'Direct Line', value: 'directline' },
			{ name: 'Emulator', value: 'emulator' },
			{ name: 'Microsoft 365 Copilot', value: 'msteamscopilot' },
			{ name: 'Microsoft Teams', value: 'msteams' },
			{ name: 'Web Chat', value: 'webchat' },
		],
		description: 'Azure Bot Service channel the conversation belongs to',
	},
	{
		displayName: 'Activity ID',
		name: 'activityId',
		type: 'string',
		default: '',
		displayOptions: { show: { conversationSource: ['manual'] } },
		description:
			'Required only for Reply / Update / Delete operations against a specific activity. Leave empty for proactive sends.',
	},
];

export function resolveConversationReference(
	ctx: IExecuteFunctions,
	itemIndex: number,
	item: IDataObject,
): ConversationReference {
	const source = ctx.getNodeParameter('conversationSource', itemIndex) as 'envelope' | 'manual';
	if (source === 'envelope') {
		const ref = item.conversationReference as ConversationReference | undefined;
		if (!ref) {
			throw new NodeOperationError(
				ctx.getNode(),
				'conversationReference is missing on the input item. Switch Conversation Source to "Specify Manually" or route through M365 Agent Trigger.',
				{ itemIndex },
			);
		}
		return ref;
	}
	const serviceUrl = ctx.getNodeParameter('serviceUrl', itemIndex) as string;
	const conversationId = ctx.getNodeParameter('conversationId', itemIndex) as string;
	const channelId = ctx.getNodeParameter('channelId', itemIndex) as string;
	const activityId = ctx.getNodeParameter('activityId', itemIndex, '') as string;
	if (!serviceUrl || !conversationId) {
		throw new NodeOperationError(
			ctx.getNode(),
			'Manual Conversation Source requires both Service URL and Conversation ID.',
			{ itemIndex },
		);
	}
	return {
		serviceUrl,
		conversation: { id: conversationId },
		channelId,
		...(activityId ? { activityId } : {}),
	};
}
