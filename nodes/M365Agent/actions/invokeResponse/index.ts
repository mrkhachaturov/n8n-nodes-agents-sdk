import type { INodeProperties } from 'n8n-workflow';
import * as respond from './respond.operation';

const resourceDisplayOptions = { show: { resource: ['invokeResponse'] } };

export const description: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: resourceDisplayOptions,
		options: [{ name: 'Respond', value: 'respond', action: 'Respond to an invoke activity' }],
		default: 'respond',
	},

	// Simple vs Advanced toggle (manifest §6.3). Default (Plain) listed first
	// so the top-of-dropdown matches the default value — trumps strict
	// alphabetical ordering per manifest §6.3 intent.
	{
		displayName: 'Response Shape',
		name: 'responseShape',
		type: 'options',
		noDataExpression: true,
		displayOptions: { show: { resource: ['invokeResponse'], operation: ['respond'] } },
		options: [
			{
				name: 'Plain (Simple)',
				value: 'plain',
				description: 'Generic { status, body } invoke response for most invoke types',
			},
			{
				name: 'Adaptive Card (Advanced)',
				value: 'adaptiveCard',
				description: 'Full AdaptiveCardInvokeResponse — required by Action.Execute invokes',
			},
		],
		default: 'plain',
	},
	{
		displayName: 'Status Code',
		name: 'statusCode',
		type: 'number',
		default: 200,
		displayOptions: { show: { resource: ['invokeResponse'], operation: ['respond'] } },
	},
	{
		displayName: 'Body',
		name: 'body',
		type: 'json',
		default: '{}',
		displayOptions: { show: { resource: ['invokeResponse'], operation: ['respond'] } },
	},
	{
		displayName: 'Card Response Type',
		name: 'cardResponseType',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['invokeResponse'],
				operation: ['respond'],
				responseShape: ['adaptiveCard'],
			},
		},
		options: [
			{ name: 'Refresh Card', value: 'application/vnd.microsoft.card.adaptive' },
			{ name: 'Send Follow-up Message', value: 'application/vnd.microsoft.activity.message' },
		],
		default: 'application/vnd.microsoft.activity.message',
	},

	...respond.description,
];

export { respond };
