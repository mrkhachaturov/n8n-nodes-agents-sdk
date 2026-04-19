/* eslint-disable @n8n/community-nodes/node-usable-as-tool --
 * Triggers are not AI-invocable tools per manifest §10.3 — a trigger fires
 * from an external webhook, it cannot be called from an agent's reasoning loop.
 * The lint rule's autofix would add `usableAsTool: true`, which is incorrect
 * for trigger nodes. Suppressing this rule keeps that intent explicit.
 */
import type {
	IWebhookFunctions,
	INodeType,
	INodeTypeDescription,
	IWebhookResponseData,
	IDataObject,
	JsonObject,
} from 'n8n-workflow';
import { NodeConnectionTypes, NodeApiError } from 'n8n-workflow';
import type { Activity } from '@microsoft/agents-activity';
import {
	activityToConversationReference,
	parseActivity,
	detectTokenSource,
} from '../../shared/envelope';
import type { ItemEnvelope, M365AgentCredentials, AuthContext, AuthKind } from '../../shared/types';
import { agent365CredentialTest } from '../../shared/auth/credentialTest';
import { validateInboundToken } from '../../shared/auth/router';
import { KNOWN_INVOKE_NAMES, InvokeNameLabel } from '../../shared/invokeNames';

/**
 * Build the `invokeNames` dropdown options from the shared registry,
 * sorted alphabetically by display name to satisfy n8n-nodes-base lint rules.
 */
const INVOKE_NAME_OPTIONS = [...KNOWN_INVOKE_NAMES]
	.map((value) => ({ name: InvokeNameLabel[value], value }))
	.sort((a, b) => a.name.localeCompare(b.name));

export class M365AgentTrigger implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'M365 Agent Trigger',
		name: 'm365AgentTrigger',
		icon: { light: 'file:../../icons/m365.svg', dark: 'file:../../icons/m365.dark.svg' },
		group: ['trigger'],
		version: 1,
		description:
			'Receive activities from Azure Bot Service. Validates JWT, parses Activity, exposes conversationReference + activity envelope.',
		defaults: { name: 'M365 Agent Trigger' },
		inputs: [],
		outputs: [NodeConnectionTypes.Main],
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
		webhooks: [
			{
				name: 'default',
				httpMethod: 'POST',
				responseMode: '={{$parameter["responseMode"]}}',
				path: 'messages',
			},
			{
				name: 'setup',
				httpMethod: 'GET',
				responseMode: 'onReceived',
				path: 'messages',
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
				displayName: 'Response Mode',
				name: 'responseMode',
				type: 'options',
				noDataExpression: true,
				options: [
					{
						name: 'Immediate (onReceived)',
						value: 'onReceived',
						description:
							'Reply 200 right away. Use for message activities where no synchronous response is needed.',
					},
					{
						name: 'Wait For Response Node (responseNode)',
						value: 'responseNode',
						description:
							'Keep the HTTP connection open until M365 Agent (resource: Invoke Response) writes to it. Required for Action.Execute and messaging-extension invokes.',
					},
				],
				default: 'onReceived',
				description:
					'How the trigger replies to the incoming HTTP request. Most bots use Immediate; invoke activities require Wait For Response Node.',
			},
			{
				displayName: 'Activity Types',
				name: 'activityTypes',
				type: 'multiOptions',
				noDataExpression: true,
				default: [],
				description:
					'Filter inbound activities by type. Empty = accept all. Populated from the full ActivityTypes enum (18 members).',
				options: [
					{ name: 'Command', value: 'command', description: 'Command invocation' },
					{
						name: 'Command Result',
						value: 'commandResult',
						description: 'Result of a command invocation',
					},
					{
						name: 'Contact Relation Update',
						value: 'contactRelationUpdate',
						description: 'Contact relationship change',
					},
					{
						name: 'Conversation Update',
						value: 'conversationUpdate',
						description: 'Members added or removed from conversation',
					},
					{
						name: 'Delete User Data',
						value: 'deleteUserData',
						description: 'User requested data deletion',
					},
					{
						name: 'End Of Conversation',
						value: 'endOfConversation',
						description: 'Conversation ended',
					},
					{ name: 'Event', value: 'event', description: 'Generic event' },
					{
						name: 'Handoff',
						value: 'handoff',
						description: 'Bot-to-bot or bot-to-human handoff',
					},
					{
						name: 'Installation Update',
						value: 'installationUpdate',
						description: 'App installed / uninstalled in a conversation',
					},
					{
						name: 'Invoke',
						value: 'invoke',
						description: 'Synchronous request expecting a response (cards, tasks)',
					},
					{
						name: 'Invoke Response',
						value: 'invokeResponse',
						description: 'Response payload for an invoke (rare inbound)',
					},
					{ name: 'Message', value: 'message', description: 'Standard text or rich message' },
					{
						name: 'Message Delete',
						value: 'messageDelete',
						description: 'A previously-sent message was deleted',
					},
					{
						name: 'Message Reaction',
						value: 'messageReaction',
						description: 'Reaction added to or removed from a message',
					},
					{
						name: 'Message Update',
						value: 'messageUpdate',
						description: 'A previously-sent message was edited',
					},
					{
						name: 'Suggestion',
						value: 'suggestion',
						description: 'Proactive suggestion activity',
					},
					{ name: 'Trace', value: 'trace', description: 'Debug / trace activity' },
					{ name: 'Typing', value: 'typing', description: 'Typing indicator signal' },
				],
			},
			{
				displayName: 'Invoke Names',
				name: 'invokeNames',
				type: 'multiOptions',
				noDataExpression: true,
				default: [],
				displayOptions: { show: { activityTypes: ['invoke'] } },
				description:
					'Filter inbound invoke activities by name. Empty = accept all invoke names. Select from known names or add custom via expression for unknown/future names.',
				options: INVOKE_NAME_OPTIONS,
			},
			{
				displayName: 'Channel Filter',
				name: 'channelFilter',
				type: 'multiOptions',
				options: [
					{ name: 'Direct Line', value: 'directline' },
					{ name: 'Emulator', value: 'emulator' },
					{ name: 'Microsoft 365 Copilot', value: 'msteamscopilot' },
					{ name: 'Microsoft Teams', value: 'msteams' },
					{ name: 'Web Chat', value: 'webchat' },
				],
				default: [],
				description: 'If set, only emit activities from these channels. Empty = all.',
			},
		],
	};

	methods = {
		credentialTest: { agent365CredentialTest },
	};

	async webhook(this: IWebhookFunctions): Promise<IWebhookResponseData> {
		const req = this.getRequestObject();

		// GET → unauthenticated health check.
		if (req.method === 'GET') {
			const res = this.getResponseObject();
			res.status(200).json({ status: 'ok', service: 'M365 Agent' });
			return { noWebhookResponse: true };
		}

		const authKind = this.getNodeParameter('authKind', 'classicBot') as string;
		const credName = authKind === 'classicBot' ? 'm365AgentApi' : 'm365Agent365Api';
		const credentials = (await this.getCredentials(credName)) as unknown as M365AgentCredentials;

		// POST → JWT validation unless explicitly bypassed for Emulator (classicBot only).
		const anonymousAllowed = credentials.anonymousAllowed;
		let validatedClaims: Record<string, unknown> | undefined;
		if (!anonymousAllowed) {
			const authHeader = req.headers.authorization as string | undefined;
			if (!authHeader || !authHeader.toLowerCase().startsWith('bearer ')) {
				const res = this.getResponseObject();
				res.status(401).json({ error: 'Missing bearer token' });
				return { noWebhookResponse: true };
			}
			try {
				const result = await validateInboundToken(
					authKind as AuthKind,
					credentials as M365AgentCredentials,
					authHeader,
				);
				validatedClaims = result.claims;
			} catch (err) {
				const res = this.getResponseObject();
				res.status(401).json({ error: 'JWT verification failed', detail: (err as Error).message });
				return { noWebhookResponse: true };
			}
		}

		const body = this.getBodyData() as unknown as Activity;
		const activityTypes = this.getNodeParameter('activityTypes', []) as string[];
		const channelFilter = this.getNodeParameter('channelFilter', []) as string[];

		// Filters return 200 OK to Azure Bot Service (acknowledging receipt)
		// but emit zero items, so the workflow short-circuits. Azure will
		// not retry; the activity is effectively swallowed on purpose.
		if (activityTypes.length > 0 && body.type && !activityTypes.includes(body.type)) {
			return { webhookResponse: { status: 200 }, workflowData: [[]] };
		}
		if (channelFilter.length > 0 && body.channelId && !channelFilter.includes(body.channelId)) {
			return { webhookResponse: { status: 200 }, workflowData: [[]] };
		}

		// Invoke-name filter: only applies when activityTypes includes 'invoke'
		// and the inbound activity is itself an invoke. Empty list = accept all
		// invoke names. Non-empty + no match ⇒ swallow (200 OK, zero items).
		if (body.type === 'invoke' && activityTypes.includes('invoke')) {
			const invokeNames = this.getNodeParameter('invokeNames', []) as string[];
			if (invokeNames.length > 0 && (!body.name || !invokeNames.includes(body.name))) {
				return { webhookResponse: { status: 200 }, workflowData: [[]] };
			}
		}

		const authHeader = req.headers.authorization as string | undefined;
		let envelope: ItemEnvelope;
		try {
			const base: ItemEnvelope = {
				conversationReference: activityToConversationReference(body),
				activity: body,
				parsed: parseActivity(body),
				raw: body,
			};
			if (body.type === 'invoke' && typeof body.name === 'string') {
				base.invokeName = body.name;
			}
			if (authKind === 'agent365' && validatedClaims && authHeader) {
				const authContext: AuthContext = {
					inboundBearer: authHeader,
					tokenSource: detectTokenSource(validatedClaims),
					validatedClaims,
				};
				base.authContext = authContext;
			}
			envelope = base;
		} catch (err) {
			const e = err as Error;
			throw new NodeApiError(this.getNode(), { message: e.message } as JsonObject, {
				message: `Malformed Activity: ${e.message}`,
			});
		}

		return {
			workflowData: [this.helpers.returnJsonArray([envelope as unknown as IDataObject])],
		};
	}
}
