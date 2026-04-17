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
import { verifyJwt } from '../../shared/verifyJwt';
import { activityToConversationReference, parseActivity } from '../../shared/envelope';
import type { ItemEnvelope, M365AgentCredentials } from '../../shared/types';

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
		usableAsTool: true,
		credentials: [{ name: 'm365AgentApi', required: true }],
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
				options: [
					{ name: 'Conversation Update', value: 'conversationUpdate' },
					{ name: 'Event', value: 'event' },
					{ name: 'Installation Update', value: 'installationUpdate' },
					{ name: 'Invoke', value: 'invoke' },
					{ name: 'Message', value: 'message' },
					{ name: 'Message Reaction', value: 'messageReaction' },
					{ name: 'Typing', value: 'typing' },
				],
				default: ['message'],
				description: 'Only emit activities of these types',
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

	async webhook(this: IWebhookFunctions): Promise<IWebhookResponseData> {
		const req = this.getRequestObject();

		// GET → unauthenticated health check.
		if (req.method === 'GET') {
			const res = this.getResponseObject();
			res.status(200).json({ status: 'ok', service: 'M365 Agent' });
			return { noWebhookResponse: true };
		}

		const credentials = (await this.getCredentials(
			'm365AgentApi',
		)) as unknown as M365AgentCredentials;

		// POST → JWT validation unless explicitly bypassed for Emulator.
		if (!credentials.anonymousAllowed) {
			const authHeader = req.headers.authorization as string | undefined;
			if (!authHeader || !authHeader.toLowerCase().startsWith('bearer ')) {
				const res = this.getResponseObject();
				res.status(401).json({ error: 'Missing bearer token' });
				return { noWebhookResponse: true };
			}
			const token = authHeader.slice(7).trim();
			try {
				await verifyJwt(token, {
					clientId: credentials.clientId,
					tenantId: credentials.tenantId,
				});
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

		let envelope: ItemEnvelope;
		try {
			envelope = {
				conversationReference: activityToConversationReference(body),
				activity: body,
				parsed: parseActivity(body),
				raw: body,
			};
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
