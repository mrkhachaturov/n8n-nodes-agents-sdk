import { updateDisplayOptions } from 'n8n-workflow';
import type { INodeProperties } from 'n8n-workflow';
import * as send from './send.operation';
import * as update from './update.operation';
import { conversationReferenceProperties } from '../../descriptions/conversationReference';

const resourceDisplayOptions = { show: { resource: ['adaptiveCard'] } };

export const description: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: resourceDisplayOptions,
		options: [
			{ name: 'Send', value: 'send', action: 'Send an adaptive card' },
			{ name: 'Update', value: 'update', action: 'Update an adaptive card' },
		],
		default: 'send',
	},

	...updateDisplayOptions(resourceDisplayOptions, conversationReferenceProperties),

	{
		displayName: 'Card Template (JSON)',
		name: 'cardTemplate',
		type: 'json',
		typeOptions: { rows: 10 },
		required: true,
		default:
			'{\n  "type": "AdaptiveCard",\n  "version": "1.4",\n  "body": [\n    { "type": "TextBlock", "text": "${title}" }\n  ]\n}',
		description:
			'Adaptive Card JSON with ${field} placeholders. Author at https://adaptivecards.microsoft.com/designer.',
		displayOptions: { show: { resource: ['adaptiveCard'], operation: ['send', 'update'] } },
	},
	{
		displayName: 'Binding Data',
		name: 'bindingData',
		type: 'json',
		default: '={{ $json }}',
		description: 'Object used to expand ${field} placeholders. Defaults to the full item.',
		displayOptions: { show: { resource: ['adaptiveCard'], operation: ['send', 'update'] } },
	},
	{
		displayName: 'Options',
		name: 'options',
		type: 'collection',
		placeholder: 'Add option',
		default: {},
		displayOptions: { show: { resource: ['adaptiveCard'], operation: ['send', 'update'] } },
		options: [
			{
				displayName: 'Fallback Text',
				name: 'fallbackText',
				type: 'string',
				default: '',
				description: 'Shown by clients that cannot render Adaptive Cards',
			},
		],
	},

	...send.description,
	...update.description,
];

export { send, update };
