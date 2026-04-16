import type {
	IDataObject,
	IExecuteFunctions,
	INodeExecutionData,
	INodeType,
	INodeTypeDescription,
	JsonObject,
} from 'n8n-workflow';
import { NodeConnectionTypes, NodeOperationError, NodeApiError } from 'n8n-workflow';
import type { Activity } from '@microsoft/agents-activity';
import { buildAuthConfig } from '../../shared/buildAuthConfig';
import { createConnector, replyInThread, type BotConnectorBundle } from '../../shared/botConnector';
import type {
	ConversationReference,
	M365AgentCredentials,
	Operation,
} from '../../shared/types';

export class M365SendActivity implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'M365 Send Activity',
		name: 'm365SendActivity',
		icon: { light: 'file:../../icons/m365.svg', dark: 'file:../../icons/m365.dark.svg' },
		group: ['output'],
		version: 1,
		description:
			'Send a message, card, or thread reply; or update/delete an existing Activity. Uses the envelope from M365AgentTrigger.',
		defaults: { name: 'M365 Send Activity' },
		inputs: [NodeConnectionTypes.Main],
		outputs: [NodeConnectionTypes.Main],
		usableAsTool: true,
		credentials: [{ name: 'm365AgentApi', required: true }],
		properties: [
			{
				displayName: 'Operation',
				name: 'operation',
				type: 'options',
				noDataExpression: true,
				// Options alphabetized by display name to satisfy n8n-nodes-base lint rule
				options: [
					{
						name: 'Delete',
						value: 'delete',
						description: 'DELETE activities/{activityId}',
						action: 'Delete activities activity id',
					},
					{
						name: 'Proactive',
						value: 'proactive',
						description: 'POST a new Activity to the conversation',
						action: 'Post a new activity to the conversation',
					},
					{
						name: 'Reply',
						value: 'reply',
						description: 'POST to activities/{activityId}',
						action: 'Post to activities activity id',
					},
					{
						name: 'Reply In Thread',
						value: 'replyInThread',
						description: 'Teams-specific: POST into a thread via ;messageid= URL suffix',
						action: 'Teams specific post into a thread via messageid url suffix',
					},
					{
						name: 'Update',
						value: 'update',
						description: 'PUT activities/{activityId}',
						action: 'Put activities activity id',
					},
				],
				default: 'reply',
			},
			{
				displayName: 'Conversation Reference',
				name: 'conversationReference',
				type: 'json',
				default: '={{ $json.conversationReference }}',
				description:
					"Routing fields from the envelope. Defaults to the item's conversationReference.",
			},
			{
				displayName: 'Activity',
				name: 'activity',
				type: 'json',
				default: '={{ $json.activity }}',
				displayOptions: { hide: { operation: ['delete'] } },
				description: 'Activity body from the envelope (typically built by M365TextMessage or M365CardTemplate)',
			},
			{
				displayName: 'Parent Activity ID',
				name: 'parentActivityId',
				type: 'string',
				default: '',
				displayOptions: { show: { operation: ['replyInThread'] } },
				description:
					'ID of the card or post that spawned the thread. Usually stored earlier in workflow state, not the inbound activityId.',
			},
		],
	};

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		// NOTE on Activity typing: the SDK's ConnectorClient methods expect
		// `Activity` but the Bot Connector wire protocol happily accepts
		// partials (type+text or type+attachments is enough for most sends).
		// Builders upstream emit Partial<Activity>; the `as Activity` casts
		// below are load-bearing only at the TS level.
		const items = this.getInputData();
		const out: INodeExecutionData[] = [];

		const credentials = (await this.getCredentials(
			'm365AgentApi',
		)) as unknown as M365AgentCredentials;
		const authConfig = buildAuthConfig(credentials);

		// Connector bundle is memoized per serviceUrl within this execute() call.
		// MSAL's internal token cache is per-ConfidentialClientApplication instance,
		// so reusing the bundle across items with the same serviceUrl avoids
		// redundant token fetches (one acquisition per unique tenant endpoint).
		const connectors = new Map<string, BotConnectorBundle>();
		const getBundle = async (serviceUrl: string): Promise<BotConnectorBundle> => {
			let b = connectors.get(serviceUrl);
			if (!b) {
				b = await createConnector(authConfig, serviceUrl);
				connectors.set(serviceUrl, b);
			}
			return b;
		};

		for (let i = 0; i < items.length; i++) {
			const operation = this.getNodeParameter('operation', i) as Operation;
			const refParam = this.getNodeParameter('conversationReference', i) as unknown;
			const ref: ConversationReference =
				typeof refParam === 'string'
					? (JSON.parse(refParam) as ConversationReference)
					: (refParam as ConversationReference);

			if (!ref?.serviceUrl || !ref.conversation?.id) {
				throw new NodeOperationError(
					this.getNode(),
					`Operation "${operation}" requires serviceUrl and conversation.id in conversationReference.`,
					{ itemIndex: i },
				);
			}

			const needsActivityId =
				operation === 'reply' || operation === 'update' || operation === 'delete';
			if (needsActivityId && !ref.activityId) {
				throw new NodeOperationError(
					this.getNode(),
					`Operation "${operation}" requires conversationReference.activityId.`,
					{ itemIndex: i },
				);
			}

			let activity: Partial<Activity> | undefined;
			if (operation !== 'delete') {
				const actParam = this.getNodeParameter('activity', i) as unknown;
				activity =
					typeof actParam === 'string'
						? (JSON.parse(actParam) as Partial<Activity>)
						: (actParam as Partial<Activity>);
				if (!activity) {
					throw new NodeOperationError(
						this.getNode(),
						`Operation "${operation}" requires an activity body.`,
						{ itemIndex: i },
					);
				}
			}

			const bundle = await getBundle(ref.serviceUrl);

			try {
				let result: IDataObject | undefined;
				switch (operation) {
					case 'reply': {
						// activityId presence is enforced by the needsActivityId guard above.
						const r = await bundle.client.replyToActivity(
							ref.conversation.id,
							ref.activityId as string,
							activity as Activity,
						);
						result = { id: r.id };
						break;
					}
					case 'proactive': {
						const r = await bundle.client.sendToConversation(
							ref.conversation.id,
							activity as Activity,
						);
						result = { id: r.id };
						break;
					}
					case 'update': {
						const r = await bundle.client.updateActivity(
							ref.conversation.id,
							ref.activityId as string,
							activity as Activity,
						);
						result = { id: r.id };
						break;
					}
					case 'delete': {
						await bundle.client.deleteActivity(ref.conversation.id, ref.activityId as string);
						result = { ok: true, deletedActivityId: ref.activityId };
						break;
					}
					case 'replyInThread': {
						const parent = this.getNodeParameter('parentActivityId', i) as string;
						if (!parent) {
							throw new NodeOperationError(
								this.getNode(),
								'replyInThread requires parentActivityId.',
								{ itemIndex: i },
							);
						}
						const r = await replyInThread(
							bundle,
							ref.conversation.id,
							parent,
							activity as Activity,
						);
						result = { id: r.id };
						break;
					}
					default: {
						throw new NodeOperationError(
							this.getNode(),
							`Unknown operation: ${String(operation)}`,
							{ itemIndex: i },
						);
					}
				}
				out.push({
					json: { ...(items[i].json as IDataObject), sendResult: result! } as IDataObject,
					pairedItem: i,
				});
			} catch (err) {
				const e = err as Error;
				throw new NodeApiError(this.getNode(), { message: e.message } as JsonObject, {
					message: `M365SendActivity ${operation} failed: ${e.message}`,
				});
			}
		}

		return [out];
	}
}
