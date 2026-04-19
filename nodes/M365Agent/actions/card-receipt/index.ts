import { updateDisplayOptions } from 'n8n-workflow';
import type { INodeProperties } from 'n8n-workflow';
import * as send from './send.operation';
import * as update from './update.operation';
import { conversationReferenceProperties } from '../../descriptions/conversationReference';

const resourceDisplayOptions = { show: { resource: ['receiptCard'] } };
const operationDisplayOptions = {
	show: { resource: ['receiptCard'], operation: ['send', 'update'] },
};

export const description: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: resourceDisplayOptions,
		options: [
			{ name: 'Send', value: 'send', action: 'Send a receipt card' },
			{ name: 'Update', value: 'update', action: 'Update a receipt card' },
		],
		default: 'send',
	},

	...updateDisplayOptions(resourceDisplayOptions, conversationReferenceProperties),

	{
		displayName: 'Card Content (JSON)',
		name: 'cardContent',
		type: 'json',
		typeOptions: { rows: 10 },
		required: true,
		default: '={}',
		description:
			'Full ReceiptCard object — title, facts, items, total, tax, vat, tap, buttons. See MS Learn cards reference.',
		displayOptions: operationDisplayOptions,
	},
	{
		displayName: 'Options',
		name: 'options',
		type: 'collection',
		placeholder: 'Add option',
		default: {},
		displayOptions: operationDisplayOptions,
		options: [
			{
				displayName: 'Fallback Text',
				name: 'fallbackText',
				type: 'string',
				default: '',
			},
		],
	},

	...send.description,
	...update.description,
];

export { send, update };
