import type { INodeTypeDescription } from 'n8n-workflow';
import { NodeConnectionTypes } from 'n8n-workflow';

export const versionDescription: INodeTypeDescription = {
	displayName: 'M365 Conversation Ref',
	name: 'm365ConversationRef',
	icon: { light: 'file:../../icons/m365.svg', dark: 'file:../../icons/m365.dark.svg' },
	group: ['transform'],
	version: 1,
	description:
		'Serialize / deserialize M365 conversation reference + auth context for persistence via n8n data nodes. NOT a storage backend — pipe the output to Redis/MySQL/DataTable or load from one.',
	defaults: { name: 'M365 Conversation Ref' },
	inputs: [NodeConnectionTypes.Main],
	outputs: [NodeConnectionTypes.Main],
	// usableAsTool intentionally OMITTED — this node serializes conversation
	// references with optional inbound bearer tokens (sensitive auth material)
	// and must not be invocable from an AI Agent's reasoning loop. n8n-workflow
	// types only accept `true | UsableAsToolDescription | undefined`, so the
	// idiomatic "off" is omission. The class-level eslint-disable in
	// M365ConversationRef.node.ts keeps that intent explicit.
	properties: [
		{
			displayName: 'Resource',
			name: 'resource',
			type: 'options',
			noDataExpression: true,
			options: [{ name: 'Serialize', value: 'serialize' }],
			default: 'serialize',
		},
		{
			displayName: 'Operation',
			name: 'operation',
			type: 'options',
			noDataExpression: true,
			displayOptions: { show: { resource: ['serialize'] } },
			options: [
				{
					name: 'Save',
					value: 'save',
					description: 'Flatten envelope fields into a JSON payload for downstream storage',
					action: 'Serialize envelope into a JSON payload',
				},
				{
					name: 'Load',
					value: 'load',
					description: 'Reconstruct envelope fields from a previously-saved payload',
					action: 'Load envelope from a previously saved payload',
				},
			],
			default: 'save',
		},
		{
			displayName: 'Include Inbound Bearer',
			name: 'includeBearer',
			type: 'boolean',
			default: false,
			displayOptions: { show: { resource: ['serialize'], operation: ['save'] } },
			description:
				'Whether to include authContext.inboundBearer (an Azure AD access token) in the serialized payload. OFF by default — the payload omits the bearer. Security: if ON, the downstream storage node MUST be secured (encrypted at rest, access-controlled). Recommend a TTL bounded by the token expiry (exp claim).',
		},
		{
			displayName: 'Payload Field',
			name: 'payloadField',
			type: 'string',
			default: 'payload',
			displayOptions: { show: { resource: ['serialize'], operation: ['load'] } },
			description: 'Name of the input JSON field that holds the previously-saved payload. Default: "payload".',
		},
	],
};
