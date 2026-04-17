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
	credentials: [{ name: 'm365AgentApi', required: true }],
	properties: [
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
