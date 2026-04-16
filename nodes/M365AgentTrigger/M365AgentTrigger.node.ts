import type {
	IHookFunctions,
	IWebhookFunctions,
	INodeType,
	INodeTypeDescription,
	IWebhookResponseData,
} from 'n8n-workflow';
import { NodeConnectionTypes } from 'n8n-workflow';

export class M365AgentTrigger implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'M365 Agent Trigger',
		name: 'm365AgentTrigger',
		icon: 'file:../../icons/m365.svg',
		group: ['trigger'],
		version: 1,
		description:
			'Receives Activity POSTs from Azure Bot Service. Validates JWT and outputs a parsed turnContext.',
		defaults: {
			name: 'M365 Agent Trigger',
		},
		inputs: [],
		outputs: [NodeConnectionTypes.Main],
		credentials: [
			{
				name: 'm365AgentApi',
				required: true,
			},
		],
		webhooks: [
			{
				name: 'default',
				httpMethod: 'POST',
				responseMode: 'onReceived',
				path: 'messages',
			},
		],
		properties: [
			{
				displayName: 'Activity Types',
				name: 'activityTypes',
				type: 'multiOptions',
				options: [
					{ name: 'Message', value: 'message' },
					{ name: 'Conversation Update', value: 'conversationUpdate' },
					{ name: 'Message Reaction', value: 'messageReaction' },
					{ name: 'Invoke', value: 'invoke' },
					{ name: 'Event', value: 'event' },
					{ name: 'Installation Update', value: 'installationUpdate' },
					{ name: 'Typing', value: 'typing' },
				],
				default: ['message'],
				description: 'Which Activity types should trigger this workflow',
			},
			{
				displayName: 'Channel Filter',
				name: 'channelFilter',
				type: 'multiOptions',
				options: [
					{ name: 'Microsoft Teams', value: 'msteams' },
					{ name: 'Microsoft 365 Copilot', value: 'msteamscopilot' },
					{ name: 'Web Chat', value: 'webchat' },
					{ name: 'Direct Line', value: 'directline' },
					{ name: 'Emulator', value: 'emulator' },
					{ name: 'Slack', value: 'slack' },
				],
				default: [],
				description: 'Only trigger for these channels. Empty = all channels.',
			},
		],
	};

	async webhook(this: IWebhookFunctions): Promise<IWebhookResponseData> {
		// TODO: JWT validation against credentials (use @microsoft/agents-hosting CloudAdapter)
		// TODO: filter by activityTypes and channelFilter
		// TODO: shape output as { activity, turnContext, channelData } typed payload
		const body = this.getBodyData();

		return {
			workflowData: [this.helpers.returnJsonArray([body as object])],
		};
	}
}
