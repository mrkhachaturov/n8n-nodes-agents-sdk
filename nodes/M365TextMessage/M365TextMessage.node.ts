import type {
	IExecuteFunctions,
	INodeExecutionData,
	INodeType,
	INodeTypeDescription,
	IDataObject,
} from 'n8n-workflow';
import { NodeConnectionTypes } from 'n8n-workflow';
import type { Activity } from '@microsoft/agents-activity';

export class M365TextMessage implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'M365 Text Message',
		name: 'm365TextMessage',
		icon: { light: 'file:../../icons/m365.svg', dark: 'file:../../icons/m365.dark.svg' },
		group: ['transform'],
		version: 1,
		description:
			'Build a text Activity and place it in the envelope. Preserves conversationReference.',
		defaults: { name: 'M365 Text Message' },
		inputs: [NodeConnectionTypes.Main],
		outputs: [NodeConnectionTypes.Main],
		usableAsTool: true,
		properties: [
			{
				displayName: 'Text',
				name: 'text',
				type: 'string',
				typeOptions: { rows: 3 },
				default: '',
				required: true,
				placeholder: 'Hello, {{ $json.parsed.userName }}!',
				description:
					'Plain text content for the message. For @mentions, use M365TeamsMention (M1).',
			},
		],
	};

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();
		const out: INodeExecutionData[] = [];

		for (let i = 0; i < items.length; i++) {
			const text = this.getNodeParameter('text', i, '') as string;

			const activity: Partial<Activity> = {
				type: 'message',
				text,
			};

			// Preserve conversationReference AND all ancillary fields (state
			// carried by the workflow through earlier nodes — e.g. threadMessage,
			// teamsMessageId, action outcomes). Only `activity` is overwritten.
			// Earlier versions dropped these, breaking downstream expressions
			// that referenced upstream state.
			const outItem: IDataObject = {
				...(items[i].json as IDataObject),
				activity: activity as unknown as IDataObject,
			};

			out.push({ json: outItem, pairedItem: i });
		}

		return [out];
	}
}
