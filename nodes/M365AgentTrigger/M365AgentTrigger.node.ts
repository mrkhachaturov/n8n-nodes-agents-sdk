import type {
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
		icon: { light: 'file:../../icons/m365.svg', dark: 'file:../../icons/m365.dark.svg' },
		group: ['trigger'],
		version: 1,
		description:
			'Receives Activity POSTs from Azure Bot Service. Validates JWT and outputs a parsed turnContext.',
		defaults: {
			name: 'M365 Agent Trigger',
		},
		inputs: [],
		outputs: [NodeConnectionTypes.Main],
		usableAsTool: true,
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
					{ name: 'Conversation Update', value: 'conversationUpdate' },
					{ name: 'Event', value: 'event' },
					{ name: 'Installation Update', value: 'installationUpdate' },
					{ name: 'Invoke', value: 'invoke' },
					{ name: 'Message', value: 'message' },
					{ name: 'Message Reaction', value: 'messageReaction' },
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
					{ name: 'Direct Line', value: 'directline' },
					{ name: 'Emulator', value: 'emulator' },
					{ name: 'Microsoft 365 Copilot', value: 'msteamscopilot' },
					{ name: 'Microsoft Teams', value: 'msteams' },
					{ name: 'Slack', value: 'slack' },
					{ name: 'Web Chat', value: 'webchat' },
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
			workflowData: [this.helpers.returnJsonArray([body])],
		};
	}
}
