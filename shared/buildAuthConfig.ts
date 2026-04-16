import type { AuthConfiguration } from '@microsoft/agents-hosting';
import type { M365AgentCredentials } from './types';

/**
 * Map an n8n credential record to the SDK's AuthConfiguration shape.
 * Enforces per-appType requirements so the MSAL layer never sees an invalid
 * combination.
 */
export function buildAuthConfig(creds: M365AgentCredentials): AuthConfiguration {
	if (!creds.clientId) {
		throw new Error('M365AgentApi: clientId is required');
	}

	const common: Pick<AuthConfiguration, 'clientId' | 'authority' | 'tenantId' | 'clientSecret'> = {
		clientId: creds.clientId,
		authority: 'https://login.microsoftonline.com',
	};

	switch (creds.appType) {
		case 'SingleTenant':
			if (!creds.tenantId) {
				throw new Error('M365AgentApi: tenantId is required for SingleTenant');
			}
			if (!creds.clientSecret) {
				throw new Error('M365AgentApi: clientSecret is required for SingleTenant');
			}
			return { ...common, tenantId: creds.tenantId, clientSecret: creds.clientSecret };

		case 'MultiTenant':
			if (!creds.clientSecret) {
				throw new Error('M365AgentApi: clientSecret is required for MultiTenant');
			}
			// No tenantId → SDK uses "botframework.com".
			return { ...common, clientSecret: creds.clientSecret };

		case 'UserAssignedMsi':
			if (!creds.tenantId) {
				throw new Error('M365AgentApi: tenantId is required for UserAssignedMsi');
			}
			// No clientSecret → SDK uses ManagedIdentityApplication.
			return { ...common, tenantId: creds.tenantId };

		default: {
			const exhaustive: never = creds.appType;
			throw new Error(`M365AgentApi: unknown appType ${exhaustive}`);
		}
	}
}
