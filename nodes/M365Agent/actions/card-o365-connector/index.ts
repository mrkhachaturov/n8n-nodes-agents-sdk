import { updateDisplayOptions } from 'n8n-workflow';
import type { INodeProperties } from 'n8n-workflow';
import * as send from './send.operation';
import * as update from './update.operation';
import { conversationReferenceProperties } from '../../descriptions/conversationReference';
import { rawActivityOverrideField } from '../../../../shared/rawActivityOverride';

const resourceDisplayOptions = { show: { resource: ['o365ConnectorCard'] } };
const operationDisplayOptions = {
	show: { resource: ['o365ConnectorCard'], operation: ['send', 'update'] },
};

export const description: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: resourceDisplayOptions,
		options: [
			{ name: 'Send', value: 'send', action: 'Send an o365 connector card' },
			{ name: 'Update', value: 'update', action: 'Update an o365 connector card' },
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
			'Teams-specific connector card. Full O365ConnectorCard object — title, text, themeColor, summary, sections, potentialAction. See https://learn.microsoft.com/microsoftteams/platform/task-modules-and-cards/cards/cards-reference#office-365-connector-card',
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
			rawActivityOverrideField,
		],
	},

	...send.description,
	...update.description,
];

export { send, update };
