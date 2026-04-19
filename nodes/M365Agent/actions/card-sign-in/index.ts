import { updateDisplayOptions } from 'n8n-workflow';
import type { INodeProperties } from 'n8n-workflow';
import * as send from './send.operation';
import * as update from './update.operation';
import { conversationReferenceProperties } from '../../descriptions/conversationReference';

const resourceDisplayOptions = { show: { resource: ['signInCard'] } };
const operationDisplayOptions = {
	show: { resource: ['signInCard'], operation: ['send', 'update'] },
};

export const description: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: resourceDisplayOptions,
		options: [
			{ name: 'Send', value: 'send', action: 'Send a sign-in card' },
			{ name: 'Update', value: 'update', action: 'Update a sign-in card' },
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
		description: 'Label shown on the sign-in button',
		displayOptions: operationDisplayOptions,
	},
	{
		displayName: 'URL',
		name: 'url',
		type: 'string',
		required: true,
		default: '',
		placeholder: 'https://.../auth/start',
		description: 'Authentication URL the user is sent to',
		displayOptions: operationDisplayOptions,
	},
	{
		displayName: 'Text',
		name: 'text',
		type: 'string',
		typeOptions: { rows: 2 },
		default: '',
		description: 'Optional body text shown above the sign-in button',
		displayOptions: operationDisplayOptions,
	},

	...send.description,
	...update.description,
];

export { send, update };
