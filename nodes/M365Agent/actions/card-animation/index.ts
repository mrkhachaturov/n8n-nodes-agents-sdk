import { updateDisplayOptions } from 'n8n-workflow';
import type { INodeProperties } from 'n8n-workflow';
import * as send from './send.operation';
import * as update from './update.operation';
import { conversationReferenceProperties } from '../../descriptions/conversationReference';
import { mediaField } from '../../../../shared/cardBuilders/media';
import { buttonsField } from '../../../../shared/cardBuilders/buttons';

const resourceDisplayOptions = { show: { resource: ['animationCard'] } };
const operationDisplayOptions = {
	show: { resource: ['animationCard'], operation: ['send', 'update'] },
};

export const description: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: resourceDisplayOptions,
		options: [
			{ name: 'Send', value: 'send', action: 'Send an animation card' },
			{ name: 'Update', value: 'update', action: 'Update an animation card' },
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
	{
		displayName: 'Poster Image URL',
		name: 'posterImageUrl',
		type: 'string',
		default: '',
		description: 'URL of a poster / preview image shown before playback',
		displayOptions: operationDisplayOptions,
	},
	{
		displayName: 'Poster Image Alt Text',
		name: 'posterImageAltText',
		type: 'string',
		default: '',
		displayOptions: operationDisplayOptions,
	},
	{ ...mediaField(), displayOptions: operationDisplayOptions },
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
				displayName: 'Aspect Ratio',
				name: 'aspect',
				type: 'string',
				default: '',
				placeholder: '16:9',
			},
			{ displayName: 'Autoloop', name: 'autoloop', type: 'boolean', default: false },
			{ displayName: 'Autostart', name: 'autostart', type: 'boolean', default: false },
			{
				displayName: 'Duration',
				name: 'duration',
				type: 'string',
				default: '',
				placeholder: 'ISO-8601 like PT30S',
			},
			{
				displayName: 'Fallback Text',
				name: 'fallbackText',
				type: 'string',
				default: '',
			},
			{ displayName: 'Shareable', name: 'shareable', type: 'boolean', default: true },
			{
				displayName: 'Value',
				name: 'value',
				type: 'json',
				default: '={}',
				typeOptions: { rows: 2 },
				description: 'Opaque object passed through as AnimationCard.value',
			},
		],
	},

	...send.description,
	...update.description,
];

export { send, update };
