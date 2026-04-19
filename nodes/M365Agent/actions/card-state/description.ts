import { updateDisplayOptions } from 'n8n-workflow';
import type { INodeProperties } from 'n8n-workflow';
import * as select from './select.operation';
import { conversationReferenceProperties } from '../../descriptions/conversationReference';
import { rawActivityOverrideField } from '../../../../shared/rawActivityOverride';

const resourceDisplayOptions = { show: { resource: ['cardState'] } };

export const description: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: resourceDisplayOptions,
		options: [
			{
				name: 'Select',
				value: 'select',
				action: 'Select a card variant by state key and send',
			},
		],
		default: 'select',
	},

	...updateDisplayOptions(resourceDisplayOptions, conversationReferenceProperties),

	{
		displayName: 'State Key',
		name: 'stateKey',
		type: 'string',
		default: '',
		required: true,
		displayOptions: { show: { resource: ['cardState'], operation: ['select'] } },
		description:
			'The state key used to pick a variant. Matched against variants[].key.',
	},
	{
		displayName: 'Variants',
		name: 'variants',
		type: 'fixedCollection',
		typeOptions: { multipleValues: true },
		default: {},
		placeholder: 'Add Variant',
		required: true,
		displayOptions: { show: { resource: ['cardState'], operation: ['select'] } },
		options: [
			{
				name: 'values',
				displayName: 'Variant',
				values: [
					{
						displayName: 'Key',
						name: 'key',
						type: 'string',
						default: '',
						required: true,
						description: 'Match value for State Key',
					},
					{
						displayName: 'Card JSON',
						name: 'cardJson',
						type: 'string',
						typeOptions: { rows: 6 },
						default: '',
						required: true,
						description:
							'Adaptive Card JSON. May contain ${placeholder} tokens expanded against Binding Data.',
					},
					{
						displayName: 'Binding Data (JSON)',
						name: 'bindingData',
						type: 'string',
						typeOptions: { rows: 2 },
						default: '',
						description:
							'Optional JSON object bound against the card via adaptivecards-templating',
					},
				],
			},
		],
	},
	{
		displayName: 'Options',
		name: 'options',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: { show: { resource: ['cardState'], operation: ['select'] } },
		options: [
			{
				displayName: 'Fallback Text',
				name: 'fallbackText',
				type: 'string',
				default: '',
				description: 'Shown by clients that cannot render Adaptive Cards',
			},
			rawActivityOverrideField,
		],
	},

	...select.description,
];

export { select };
