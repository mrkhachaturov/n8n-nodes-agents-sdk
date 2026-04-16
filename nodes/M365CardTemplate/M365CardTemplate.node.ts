import type {
	IExecuteFunctions,
	INodeExecutionData,
	INodeType,
	INodeTypeDescription,
	IDataObject,
} from 'n8n-workflow';
import { NodeConnectionTypes, NodeOperationError } from 'n8n-workflow';
import { Template } from 'adaptivecards-templating';
import type { Activity } from '@microsoft/agents-activity';
import { mergeEnvelope } from '../../shared/envelope';
import type { ItemEnvelope } from '../../shared/types';

export class M365CardTemplate implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'M365 Card Template',
		name: 'm365CardTemplate',
		icon: { light: 'file:../../icons/m365.svg', dark: 'file:../../icons/m365.dark.svg' },
		group: ['transform'],
		version: 1,
		description:
			'Adaptive Card with ${field} templating. Paste a template from the Adaptive Cards Designer; bind data; get a finished card in the envelope.',
		defaults: { name: 'M365 Card Template' },
		inputs: [NodeConnectionTypes.Main],
		outputs: [NodeConnectionTypes.Main],
		usableAsTool: true,
		properties: [
			{
				displayName: 'Card Template (JSON)',
				name: 'cardTemplate',
				type: 'json',
				default:
					'{\n  "type": "AdaptiveCard",\n  "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",\n  "version": "1.4",\n  "body": [\n    { "type": "TextBlock", "text": "${title}", "weight": "bolder", "size": "large" }\n  ]\n}',
				required: true,
				description:
					'Adaptive Card JSON with ${field} placeholders. Author in https://adaptivecards.io/designer/.',
			},
			{
				displayName: 'Binding Data',
				name: 'bindingData',
				type: 'json',
				default: '={{ $json }}',
				description: 'Object used to expand ${field} placeholders. Defaults to the full item JSON, so fields like $JSON.order_number are available as ${order_number}. Accepts a JSON object or expression.',
			},
		],
	};

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();
		const out: INodeExecutionData[] = [];

		for (let i = 0; i < items.length; i++) {
			const current = items[i].json as unknown as Partial<ItemEnvelope>;

			const templateParam = this.getNodeParameter('cardTemplate', i) as unknown;
			const bindingDataParam = this.getNodeParameter('bindingData', i) as unknown;

			const templateObj: object =
				typeof templateParam === 'string' ? JSON.parse(templateParam) : (templateParam as object);
			const bindingObj: object =
				typeof bindingDataParam === 'string'
					? JSON.parse(bindingDataParam)
					: (bindingDataParam as object);

			let renderedCard: object;
			try {
				const tmpl = new Template(templateObj);
				renderedCard = tmpl.expand({ $root: bindingObj }) as object;
			} catch (err) {
				throw new NodeOperationError(
					this.getNode(),
					`Adaptive Card templating failed: ${(err as Error).message}`,
					{ itemIndex: i },
				);
			}

			const adaptiveCardAttachment = {
				contentType: 'application/vnd.microsoft.card.adaptive',
				content: renderedCard as unknown,
			};

			const activity: Partial<Activity> = {
				type: 'message',
				attachments: [adaptiveCardAttachment],
			};

			const outItem = current.conversationReference
				? mergeEnvelope(
						{ conversationReference: current.conversationReference, activity: current.activity },
						activity,
					)
				: { activity };

			out.push({ json: outItem as unknown as IDataObject, pairedItem: i });
		}

		return [out];
	}
}
