import { updateDisplayOptions } from 'n8n-workflow';
import type { INodeProperties } from 'n8n-workflow';
import * as send from './send.operation';
import * as reply from './reply.operation';
import * as replyInThread from './replyInThread.operation';
import * as update from './update.operation';
import * as deleteMessage from './deleteMessage.operation';
import * as typing from './typing.operation';
import { conversationReferenceProperties } from '../../descriptions/conversationReference';
import { mentionsOption } from '../../descriptions/mentionsOption';
import { suggestedActionsOption } from '../../descriptions/suggestedActionsOption';

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
			{
				name: 'Reply in Thread',
				value: 'replyInThread',
				action: 'Reply inside an existing thread',
			},
			{ name: 'Send', value: 'send', action: 'Send a message' },
			{ name: 'Typing', value: 'typing', action: 'Send typing indicator' },
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
		description: 'Message content. For @mentions, use the Options collection.',
		displayOptions: {
			show: { resource: ['message'], operation: ['send', 'reply', 'update', 'replyInThread'] },
		},
	},

	// Parent Activity ID — required top-level field for replyInThread (manifest §6.5: required fields first)
	{
		displayName: 'Parent Activity ID',
		name: 'parentActivityId',
		type: 'string',
		required: true,
		default: '',
		placeholder: 'e.g. 1673348720590',
		description:
			'ID of the card or post that spawned the thread. Usually stored earlier in the workflow (e.g. teams_message_id), NOT the inbound activityId.',
		displayOptions: { show: { resource: ['message'], operation: ['replyInThread'] } },
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
			mentionsOption,
			suggestedActionsOption,
		],
	},

	// Per-operation spreads appended last for things only one op needs (currently empty)
	...send.description,
	...reply.description,
	...replyInThread.description,
	...update.description,
	...deleteMessage.description,
	...typing.description,
];

export { send, reply, replyInThread, update, deleteMessage, typing };
