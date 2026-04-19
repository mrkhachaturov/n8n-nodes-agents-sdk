import { NodeConnectionTypes, type INodeTypeDescription } from 'n8n-workflow';
import * as message from './message';
import * as cardAdaptive from './card-adaptive';
import * as cardAnimation from './card-animation';
import * as cardAudio from './card-audio';
import * as cardHero from './card-hero';
import * as cardReceipt from './card-receipt';
import * as cardSignIn from './card-sign-in';
import * as cardThumbnail from './card-thumbnail';
import * as cardVideo from './card-video';
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
		{
			name: 'm365AgentApi',
			required: true,
			displayOptions: { show: { authKind: ['classicBot'] } },
		},
		{
			name: 'm365Agent365Api',
			required: true,
			testedBy: 'agent365CredentialTest',
			displayOptions: { show: { authKind: ['agent365'] } },
		},
	],
	properties: [
		{
			displayName: 'Authentication Kind',
			name: 'authKind',
			type: 'options',
			noDataExpression: true,
			options: [
				{
					name: 'Classic Bot (Azure Bot Service)',
					value: 'classicBot',
					description: 'Existing Azure Bot resource + Entra App Registration',
				},
				{
					name: 'Agent 365 (Entra Agent Identity)',
					value: 'agent365',
					description: 'Agent Blueprint in Entra, no Azure Bot resource needed',
				},
			],
			default: 'classicBot',
		},
		{
			displayName: 'Resource',
			name: 'resource',
			type: 'options',
			noDataExpression: true,
			options: [
				{ name: 'Adaptive Card', value: 'adaptiveCard' },
				{ name: 'Animation Card', value: 'animationCard' },
				{ name: 'Audio Card', value: 'audioCard' },
				{ name: 'Hero Card', value: 'heroCard' },
				{ name: 'Invoke Response', value: 'invokeResponse' },
				{ name: 'Message', value: 'message' },
				{ name: 'Receipt Card', value: 'receiptCard' },
				{ name: 'Sign-In Card', value: 'signInCard' },
				{ name: 'Thumbnail Card', value: 'thumbnailCard' },
				{ name: 'Video Card', value: 'videoCard' },
			],
			default: 'message',
		},
		...message.description,
		...cardAdaptive.description,
		...cardAnimation.description,
		...cardAudio.description,
		...cardHero.description,
		...cardReceipt.description,
		...cardSignIn.description,
		...cardThumbnail.description,
		...cardVideo.description,
		...invokeResponse.description,
	],
};
