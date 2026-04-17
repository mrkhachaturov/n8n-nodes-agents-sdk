import { updateDisplayOptions } from 'n8n-workflow';
import type { INodeProperties } from 'n8n-workflow';
import * as send from './send.operation';
import * as reply from './reply.operation';
import * as update from './update.operation';
import * as deleteMessage from './deleteMessage.operation';
import { conversationReferenceProperties } from '../../descriptions/conversationReference';

const resourceDisplayOptions = { show: { resource: ['message'] } };

export const description: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: resourceDisplayOptions,
		options: [
			{ name: 'Delete', value: 'delete', action: 'Delete a message' },
			{ name: 'Reply', value: 'reply', action: 'Reply to a message' },
			{ name: 'Send', value: 'send', action: 'Send a message' },
			{ name: 'Update', value: 'update', action: 'Update a message' },
		],
		default: 'send',
	},

	// Conversation reference — shown for every message operation
	...updateDisplayOptions(resourceDisplayOptions, conversationReferenceProperties),

	// Text field — hidden for delete; shown for body-carrying ops (send/reply/update) + replyInThread (Task 6)
	{
		displayName: 'Text',
		name: 'text',
		type: 'string',
		typeOptions: { rows: 3 },
		required: true,
		default: '',
		placeholder: 'Hello, {{ $json.parsed.userName }}!',
		description: 'Message content. For @mentions, use the Options collection (M1+).',
		displayOptions: {
			show: { resource: ['message'], operation: ['send', 'reply', 'update', 'replyInThread'] },
		},
	},

	// Options collection — progressive disclosure per manifest §6.4
	{
		displayName: 'Options',
		name: 'options',
		type: 'collection',
		placeholder: 'Add option',
		default: {},
		displayOptions: {
			show: { resource: ['message'], operation: ['send', 'reply', 'update', 'replyInThread'] },
		},
		options: [
			{
				displayName: 'Workflow Footer',
				name: 'workflowFooter',
				type: 'boolean',
				default: false,
				description:
					'Whether to append "_Sent from n8n workflow._" as a plain-text marker at the end of the message. Helpful for debugging and auditing. (A clickable link with instance URL + workflow ID is deferred to a later milestone.)',
			},
			// M1 will extend this with mentions, suggestedActions, etc.
		],
	},

	// Per-operation spreads appended last for things only one op needs (currently empty)
	...send.description,
	...reply.description,
	...update.description,
	...deleteMessage.description,
];

export { send, reply, update, deleteMessage };
