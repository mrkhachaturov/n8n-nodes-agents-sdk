import type {
	ICredentialTestRequest,
	ICredentialType,
	INodeProperties,
} from 'n8n-workflow';

export class M365AgentApi implements ICredentialType {
	name = 'm365AgentApi';

	displayName = 'Microsoft 365 Agent API';

	documentationUrl =
		'https://learn.microsoft.com/microsoft-365/agents-sdk/';

	icon = { light: 'file:../icons/m365.svg', dark: 'file:../icons/m365.dark.svg' } as const;

	// Credential "test" button — verifies outbound HTTPS to Entra, not
	// that the App ID / secret / tenant are valid. A green check here
	// only means n8n can reach login.microsoftonline.com. Actual
	// credential validation happens on the first M365SendActivity call.
	test: ICredentialTestRequest = {
		request: {
			baseURL: 'https://login.microsoftonline.com',
			url: '/common/v2.0/.well-known/openid-configuration',
			method: 'GET',
		},
	};

	properties: INodeProperties[] = [
		{
			displayName: 'App Type',
			name: 'appType',
			type: 'options',
			options: [
				{ name: 'Single Tenant', value: 'SingleTenant' },
				{ name: 'Multi Tenant', value: 'MultiTenant' },
				{ name: 'User-Assigned Managed Identity', value: 'UserAssignedMsi' },
			],
			default: 'SingleTenant',
			description:
				'Azure Bot registration app type. Multi-tenant creation is deprecated after 2025-07-31; prefer SingleTenant or UserAssignedMsi.',
		},
		{
			displayName: 'App ID (Client ID)',
			name: 'clientId',
			type: 'string',
			default: '',
			required: true,
			description: 'Microsoft App ID from the Azure Bot registration',
		},
		{
			displayName: 'Client Secret',
			name: 'clientSecret',
			type: 'string',
			typeOptions: { password: true },
			default: '',
			displayOptions: {
				show: { appType: ['SingleTenant', 'MultiTenant'] },
			},
			description: 'Client secret from the App Registration',
		},
		{
			displayName: 'Tenant ID',
			name: 'tenantId',
			type: 'string',
			default: '',
			displayOptions: {
				show: { appType: ['SingleTenant', 'UserAssignedMsi'] },
			},
			description: 'Microsoft Entra tenant ID',
		},
		{
			displayName: 'Allow Anonymous (Emulator)',
			name: 'anonymousAllowed',
			type: 'boolean',
			default: false,
			description:
				'Skip JWT validation. Only for Bot Framework Emulator testing. Never true in production.',
		},
	];
}
