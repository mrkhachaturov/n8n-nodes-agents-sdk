import type {
	IExecuteFunctions,
	INodeExecutionData,
	INodeType,
	INodeTypeDescription,
} from 'n8n-workflow';
import { NodeConnectionTypes } from 'n8n-workflow';

export class M365SendActivity implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'M365 Send Activity',
		name: 'm365SendActivity',
		icon: 'file:../../icons/m365.svg',
		group: ['output'],
		version: 1,
		description:
			'Send a message, Adaptive Card, or update/delete an existing Activity via Azure Bot Service.',
		defaults: {
			name: 'M365 Send Activity',
		},
		inputs: [NodeConnectionTypes.Main],
		outputs: [NodeConnectionTypes.Main],
		usableAsTool: true,
		credentials: [
			{
				name: 'm365AgentApi',
				required: true,
			},
		],
		properties: [
			{
				displayName: 'Operation',
				name: 'operation',
				type: 'options',
				noDataExpression: true,
				options: [
					{ name: 'Send Reply', value: 'reply' },
					{ name: 'Send Proactive Message', value: 'proactive' },
					{ name: 'Update Activity', value: 'update' },
					{ name: 'Delete Activity', value: 'delete' },
				],
				default: 'reply',
			},
			{
				displayName: 'Conversation Reference',
				name: 'conversationReference',
				type: 'json',
				default: '={{ $json.activity }}',
				description:
					'The Activity or ConversationReference to reply to. Defaults to the incoming activity from M365 Agent Trigger.',
			},
			{
				displayName: 'Content Type',
				name: 'contentType',
				type: 'options',
				displayOptions: {
					show: { operation: ['reply', 'proactive', 'update'] },
				},
				options: [
					{ name: 'Text', value: 'text' },
					{ name: 'Adaptive Card', value: 'adaptiveCard' },
				],
				default: 'text',
			},
			{
				displayName: 'Text',
				name: 'text',
				type: 'string',
				typeOptions: { rows: 3 },
				displayOptions: {
					show: { contentType: ['text'] },
				},
				default: '',
			},
			{
				displayName: 'Adaptive Card JSON',
				name: 'card',
				type: 'json',
				displayOptions: {
					show: { contentType: ['adaptiveCard'] },
				},
				default: '',
				description: 'Adaptive Card payload (schema 1.5+ recommended for Teams)',
			},
		],
	};

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();
		const returnData: INodeExecutionData[] = [];

		for (let i = 0; i < items.length; i++) {
			// TODO: build CloudAdapter from credentials
			// TODO: route by operation (reply / proactive / update / delete)
			// TODO: use adapter.continueConversationAsync() for proactive + update + delete
			// TODO: capture response, push activityId into returnData for chaining
			returnData.push({ json: { ok: true, stub: true } });
		}

		return [returnData];
	}
}
