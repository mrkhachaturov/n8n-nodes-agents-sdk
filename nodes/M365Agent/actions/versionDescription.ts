import { NodeConnectionTypes, type INodeTypeDescription } from 'n8n-workflow';
import * as message from './message';
import * as card from './card';
import * as invokeResponse from './invokeResponse';

export const versionDescription: INodeTypeDescription = {
	displayName: 'M365 Agent',
	name: 'm365Agent',
	icon: { light: 'file:../../icons/m365.svg', dark: 'file:../../icons/m365.dark.svg' },
	group: ['output'],
	version: 1,
	subtitle: '={{$parameter["operation"] + ": " + $parameter["resource"]}}',
	description:
		'Send messages, cards, thread replies, and invoke responses through Microsoft 365 Agents (Teams / Copilot / WebChat).',
	defaults: { name: 'M365 Agent' },
	inputs: [NodeConnectionTypes.Main],
	outputs: [NodeConnectionTypes.Main],
	// usableAsTool intentionally OMITTED — this node performs side effects
	// (sends/updates/deletes activities to Azure Bot Service) and should not be
	// invocable from an AI Agent's reasoning loop. See manifest §10.7.
	credentials: [
		{ name: 'm365AgentApi', required: true, displayOptions: { show: { authKind: ['classicBot'] } } },
		{ name: 'm365Agent365Api', required: true, displayOptions: { show: { authKind: ['agent365'] } } },
	],
	properties: [
		{
			displayName: 'Authentication Kind',
			name: 'authKind',
			type: 'options',
			noDataExpression: true,
			options: [
				{ name: 'Classic Bot (Azure Bot Service)', value: 'classicBot',
					description: 'Existing Azure Bot resource + Entra App Registration.' },
				{ name: 'Agent 365 (Entra Agent Identity)', value: 'agent365',
					description: 'Agent Blueprint in Entra, no Azure Bot resource needed.' },
			],
			default: 'classicBot',
		},
		{
			displayName: 'Resource',
			name: 'resource',
			type: 'options',
			noDataExpression: true,
			options: [
				{ name: 'Card', value: 'card' },
				{ name: 'Invoke Response', value: 'invokeResponse' },
				{ name: 'Message', value: 'message' },
			],
			default: 'message',
		},
		...message.description,
		...card.description,
		...invokeResponse.description,
	],
};
