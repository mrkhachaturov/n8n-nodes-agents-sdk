import { updateDisplayOptions } from 'n8n-workflow';
import type { INodeProperties } from 'n8n-workflow';
import * as send from './send.operation';
import * as update from './update.operation';
import { conversationReferenceProperties } from '../../descriptions/conversationReference';
import { rawActivityOverrideField } from '../../../../shared/rawActivityOverride';

const resourceDisplayOptions = { show: { resource: ['rawAttachment'] } };
const operationDisplayOptions = {
	show: { resource: ['rawAttachment'], operation: ['send', 'update'] },
};

export const description: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: resourceDisplayOptions,
		options: [
			{ name: 'Send', value: 'send', action: 'Send a raw attachment' },
			{ name: 'Update', value: 'update', action: 'Update a raw attachment' },
		],
		default: 'send',
	},

	...updateDisplayOptions(resourceDisplayOptions, conversationReferenceProperties),

	{
		displayName: 'Content Type',
		name: 'contentType',
		type: 'string',
		required: true,
		default: '',
		placeholder: 'application/vnd.microsoft.card.adaptive',
		description: 'Attachment MIME / card content type',
		displayOptions: operationDisplayOptions,
	},
	{
		displayName: 'Content (JSON)',
		name: 'content',
		type: 'json',
		typeOptions: { rows: 8 },
		required: true,
		default: '={}',
		description: 'Attachment content — the object that goes into Attachment.content',
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
			{ displayName: 'Content URL', name: 'contentUrl', type: 'string', default: '' },
			{ displayName: 'Name', name: 'name', type: 'string', default: '' },
			rawActivityOverrideField,
			{ displayName: 'Thumbnail URL', name: 'thumbnailUrl', type: 'string', default: '' },
		],
	},

	...send.description,
	...update.description,
];

export { send, update };
