import type { ICredentialType, INodeProperties } from 'n8n-workflow';

export class M365Agent365Api implements ICredentialType {
	name = 'm365Agent365Api';

	displayName = 'M365 Agent 365 API';

	documentationUrl = 'https://learn.microsoft.com/entra/agent-id/';

	icon = { light: 'file:../icons/m365.svg', dark: 'file:../icons/m365.dark.svg' } as const;

	testedBy = 'agent365CredentialTest';

	properties: INodeProperties[] = [
		{
			displayName: 'Tenant ID',
			name: 'tenantId',
			type: 'string',
			default: '',
			required: true,
			description: 'Entra tenant hosting the Blueprint.',
		},
		{
			displayName: 'Blueprint App ID',
			name: 'blueprintAppId',
			type: 'string',
			default: '',
			required: true,
			description: 'Client ID of the Agent Identity Blueprint (from `a365 config display -g`).',
		},
		{
			displayName: 'Agent Identity Instance App ID',
			name: 'agentInstanceAppId',
			type: 'string',
			default: '',
			required: true,
			displayOptions: { show: { transport: ['sidecar'] } },
			description:
				'Entra Agent Identity instance appId (distinct from Blueprint App ID). Sidecar uses this as the AgentIdentity query param.',
		},
		{
			displayName: 'Outbound Token Source',
			name: 'transport',
			type: 'options',
			noDataExpression: true,
			required: true,
			options: [
				{
					name: 'Sidecar (Entra SDK for Agent ID)',
					value: 'sidecar',
					description: 'HTTP calls to a companion container. Enables all three identity modes.',
				},
				{
					name: 'Inline (Agents SDK)',
					value: 'inline',
					description: 'In-process via @microsoft/agents-hosting. Autonomous only in M1.',
				},
			],
			default: 'sidecar',
		},
		{
			displayName: 'Blueprint Credential Type',
			name: 'inlineCredKind',
			type: 'options',
			noDataExpression: true,
			required: true,
			displayOptions: { show: { transport: ['inline'] } },
			options: [
				{ name: 'Client Secret', value: 'clientSecret' },
				{ name: 'Client Certificate (PEM)', value: 'clientCert' },
			],
			default: 'clientSecret',
		},
		{
			displayName: 'Blueprint Client Secret',
			name: 'blueprintSecret',
			type: 'string',
			typeOptions: { password: true },
			default: '',
			required: true,
			displayOptions: { show: { transport: ['inline'], inlineCredKind: ['clientSecret'] } },
		},
		{
			displayName: 'Blueprint Client Certificate (PEM)',
			name: 'blueprintCertPem',
			type: 'string',
			typeOptions: { password: true, rows: 8 },
			default: '',
			required: true,
			displayOptions: { show: { transport: ['inline'], inlineCredKind: ['clientCert'] } },
		},
		{
			displayName: 'Certificate Thumbprint',
			name: 'blueprintCertThumbprint',
			type: 'string',
			default: '',
			required: true,
			displayOptions: { show: { transport: ['inline'], inlineCredKind: ['clientCert'] } },
		},
		{
			displayName: 'Sidecar Base URL',
			name: 'sidecarUrl',
			type: 'string',
			default: 'http://entra-auth-sidecar:5000',
			required: true,
			displayOptions: { show: { transport: ['sidecar'] } },
			description: 'Docker service DNS of the sidecar on the same overlay network.',
		},
		{
			displayName: 'Outbound Downstream API Name',
			name: 'outboundDownstreamApi',
			type: 'string',
			default: 'MessagingBotApi',
			required: true,
			displayOptions: { show: { transport: ['sidecar'] } },
			description: "Matches the sidecar's DownstreamApis__<name>__* env var group.",
		},
		{
			displayName: 'Default Agent Username',
			name: 'defaultAgentUsername',
			type: 'string',
			default: '',
			description:
				'Optional default when operation picks Identity Mode = Agent User Account.',
		},
		{
			displayName: 'Inbound Token Validation',
			name: 'validateVia',
			type: 'options',
			noDataExpression: true,
			required: true,
			options: [
				{ name: 'Match Outbound Transport', value: 'sameAsOutbound' },
				{ name: 'Inline (Extended JWKS)', value: 'inline' },
				{ name: 'Sidecar /Validate', value: 'sidecar' },
			],
			default: 'sameAsOutbound',
		},
		{
			displayName: 'Validator Sidecar URL',
			name: 'validatorSidecarUrl',
			type: 'string',
			default: '',
			required: true,
			displayOptions: { show: { validateVia: ['sidecar'], transport: ['inline'] } },
			description: 'Required when validation goes sidecar but outbound is inline.',
		},
		{
			displayName: 'Options',
			name: 'options',
			type: 'collection',
			placeholder: 'Add option',
			default: {},
			options: [
				{
					displayName: 'Allowed Issuers Override',
					name: 'allowedIssuers',
					type: 'string',
					typeOptions: { rows: 3 },
					default: '',
				},
				{
					displayName: 'Allowed Audiences Override',
					name: 'allowedAudiences',
					type: 'string',
					typeOptions: { rows: 3 },
					default: '',
				},
				{
					displayName: 'Sidecar Request Timeout (ms)',
					name: 'sidecarTimeoutMs',
					type: 'number',
					default: 30000,
					typeOptions: { minValue: 1000 },
				},
				{
					displayName: 'Token Cache TTL (seconds)',
					name: 'tokenCacheTtlSec',
					type: 'number',
					default: 300,
					typeOptions: { minValue: 0 },
				},
			],
		},
	];
}
