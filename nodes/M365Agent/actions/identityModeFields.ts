import type { INodeProperties } from 'n8n-workflow';

type SupportedResource = 'message' | 'card';
type MessageOperation = 'send' | 'reply' | 'update' | 'delete' | 'replyInThread';
type CardOperation = 'send' | 'update';

/** Build identityMode + agent-user selector fields for a given resource+operation pair.
 *  Produces 5 properties all gated by authKind=agent365, the resource name, and the
 *  caller's operation name. */
export function identityModeFields(
	resource: 'message',
	operation: MessageOperation,
): INodeProperties[];
export function identityModeFields(
	resource: 'card',
	operation: CardOperation,
): INodeProperties[];
export function identityModeFields(
	resource: SupportedResource,
	operation: string,
): INodeProperties[] {
	const base: Record<string, string[]> = {
		authKind: ['agent365'],
		resource: [resource],
		operation: [operation],
	};
	const withAgentUser = { ...base, identityMode: ['agentUser' as const] };
	return [
		{
			displayName: 'Identity Mode',
			name: 'identityMode',
			type: 'options',
			noDataExpression: true,
			displayOptions: { show: base },
			options: [
				{
					name: 'Autonomous',
					value: 'autonomous',
					description: 'Agent acts as its Blueprint app identity.',
				},
				{
					name: 'Agent User Account',
					value: 'agentUser',
					description:
						'Agent acts as its own M365 user (requires sidecar transport and a separate license).',
				},
			],
			default: 'autonomous',
		},
		{
			displayName: 'User Selector Mode',
			name: 'userSelectorMode',
			type: 'options',
			noDataExpression: true,
			displayOptions: { show: withAgentUser },
			options: [
				{ name: 'By UPN', value: 'byUpn' },
				{ name: 'By Object ID', value: 'byObjectId' },
			],
			default: 'byUpn',
		},
		{
			displayName: 'Agent Username',
			name: 'agentUsername',
			type: 'string',
			default: '={{ $credentials.defaultAgentUsername || "" }}',
			placeholder: 'agent@contoso.com',
			displayOptions: { show: { ...withAgentUser, userSelectorMode: ['byUpn' as const] } },
			description:
				"UPN of the agentic user. Defaults to credential's Default Agent Username.",
		},
		{
			displayName: 'Agent User ID',
			name: 'agentUserId',
			type: 'string',
			default: '',
			placeholder: 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx',
			displayOptions: { show: { ...withAgentUser, userSelectorMode: ['byObjectId' as const] } },
			description: 'Object ID of the agentic user.',
		},
		{
			displayName:
				"Agent User Account mode uses the agent's own M365 user and requires a separate license (E5 + Teams Enterprise + Agent 365) on the agentic user account.",
			name: 'agentUserNotice',
			type: 'notice',
			default: '',
			displayOptions: { show: withAgentUser },
		},
	];
}
