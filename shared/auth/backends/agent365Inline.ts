import { MsalTokenProvider } from '@microsoft/agents-hosting';
import { NodeOperationError } from 'n8n-workflow';
import type { INode } from 'n8n-workflow';
import type { M365Agent365Cred, IdentityMode } from '../../types';
import { buildBlueprintAuthConfig, scopeForDownstream } from '../blueprintAuthConfig';

export async function acquireAgent365InlineToken(
	cred: M365Agent365Cred,
	identityMode: IdentityMode,
	downstreamApi: string,
	node?: INode,
): Promise<{ authorizationHeader: string }> {
	if (identityMode === 'agentUser') {
		throw new NodeOperationError(
			node ?? ({} as INode),
			'Agent User mode requires sidecar transport. Switch transport in the credential.',
		);
	}
	if (identityMode === 'interactiveOBO') {
		throw new NodeOperationError(
			node ?? ({} as INode),
			'Interactive OBO inline transport is not implemented in M1. Use sidecar transport, or wait for M2 Graph/MCP operations.',
		);
	}
	// autonomous
	const authConfig = buildBlueprintAuthConfig(cred);
	const provider = new MsalTokenProvider();
	const token = await provider.getAccessToken(authConfig, scopeForDownstream(downstreamApi));
	return { authorizationHeader: `Bearer ${token}` };
}
