import { updateDisplayOptions } from 'n8n-workflow';
import type { INodeProperties } from 'n8n-workflow';
import * as send from './send.operation';
import * as update from './update.operation';
import { conversationReferenceProperties } from '../../descriptions/conversationReference';
import { imagesField } from '../../../../shared/cardBuilders/images';
import { buttonsField } from '../../../../shared/cardBuilders/buttons';
import { cardActionFields } from '../../../../shared/cardBuilders/cardActionRow';
import { rawActivityOverrideField } from '../../../../shared/rawActivityOverride';

const resourceDisplayOptions = { show: { resource: ['heroCard'] } };
const operationDisplayOptions = {
	show: { resource: ['heroCard'], operation: ['send', 'update'] },
};

export const description: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: resourceDisplayOptions,
		options: [
			{ name: 'Send', value: 'send', action: 'Send a hero card' },
			{ name: 'Update', value: 'update', action: 'Update a hero card' },
		],
		default: 'send',
	},

	...updateDisplayOptions(resourceDisplayOptions, conversationReferenceProperties),

	{
		displayName: 'Title',
		name: 'title',
		type: 'string',
		required: true,
		default: '',
		displayOptions: operationDisplayOptions,
	},
	{
		displayName: 'Subtitle',
		name: 'subtitle',
		type: 'string',
		default: '',
		displayOptions: operationDisplayOptions,
	},
	{
		displayName: 'Text',
		name: 'text',
		type: 'string',
		typeOptions: { rows: 3 },
		default: '',
		displayOptions: operationDisplayOptions,
	},
	{ ...imagesField(), displayOptions: operationDisplayOptions },
	{ ...buttonsField(), displayOptions: operationDisplayOptions },
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
				description: 'Shown by clients that cannot render this card',
			},
			rawActivityOverrideField,
			{
				displayName: 'Tap Action',
				name: 'tapAction',
				type: 'fixedCollection',
				typeOptions: { multipleValues: false },
				placeholder: 'Add Tap Action',
				default: {},
				description: 'Action invoked when the card body is tapped',
				options: [
					{
						displayName: 'Action',
						name: 'action',
						values: cardActionFields(),
					},
				],
			},
		],
	},

	...send.description,
	...update.description,
];

export { send, update };
