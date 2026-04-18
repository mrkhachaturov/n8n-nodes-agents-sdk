import type {
	IDataObject, IExecuteFunctions, INodeExecutionData,
} from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import type { BotConnectorBundle } from '../../../shared/botConnector';
import type {
	AuthKind, IdentityMode, M365ClassicBotCred, M365Agent365Cred,
} from '../../../shared/types';

import * as message from './message';
import * as card from './card';
import * as invokeResponse from './invokeResponse';

type Resource = 'message' | 'card' | 'invokeResponse';

export function makeBundleKey(
	serviceUrl: string,
	authKind: AuthKind,
	identityMode: IdentityMode,
	agentUsername?: string,
	agentUserId?: string,
): string {
	return [serviceUrl, authKind, identityMode, agentUsername ?? '', agentUserId ?? ''].join('::');
}

export async function router(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
	const items = this.getInputData();
	const out: INodeExecutionData[] = [];

	// Invoke Response is a webhook-response writer (see manifest §9 — mirrors
	// RespondToWebhook). It must run ONCE for the batch, not per item, or we'd
	// call `this.sendResponse(...)` multiple times against a single HTTP request.
	// Special-case it before the per-item loop; pass items through unchanged.
	// Credentials are NOT fetched on this path — respond.execute() only writes
	// to the HTTP connection and never talks to Azure Bot Service.
	const firstResource = this.getNodeParameter('resource', 0) as Resource;
	if (firstResource === 'invokeResponse') {
		const firstOp = this.getNodeParameter('operation', 0) as string;
		if (firstOp !== 'respond') {
			throw new NodeOperationError(this.getNode(), `Unknown invokeResponse operation: ${firstOp}`, { itemIndex: 0 });
		}
		await invokeResponse.respond.execute(this);
		return [items.map((item, i) => ({ ...item, pairedItem: i }))];
	}

	// Credentials are only needed for message/card operations (both talk to the
	// Bot Connector via MSAL). Fetching after the invokeResponse early-return
	// avoids a pointless vault roundtrip on the response-only path.
	// authKind drives which credential family to fetch — classicBot uses the
	// legacy M365AgentApi credential (App ID + secret/MSI), agent365 uses the
	// new M365Agent365Api credential (Blueprint App + sidecar/inline transport).
	const authKind = this.getNodeParameter('authKind', 0, 'classicBot') as AuthKind;
	const credentialsRaw = authKind === 'classicBot'
		? await this.getCredentials('m365AgentApi')
		: await this.getCredentials('m365Agent365Api');
	const credentials = credentialsRaw as unknown as M365ClassicBotCred | M365Agent365Cred;

	const bundles = new Map<string, BotConnectorBundle>();

	for (let i = 0; i < items.length; i++) {
		const resource = this.getNodeParameter('resource', i) as Resource;
		const operation = this.getNodeParameter('operation', i) as string;

		let result: IDataObject;
		switch (resource) {
			case 'message': {
				switch (operation) {
					case 'send':
						result = await message.send.execute(this, i, authKind, credentials, bundles);
						break;
					case 'reply':
						result = await message.reply.execute(this, i, authKind, credentials, bundles);
						break;
					case 'update':
						result = await message.update.execute(this, i, authKind, credentials, bundles);
						break;
					case 'delete':
						result = await message.deleteMessage.execute(this, i, authKind, credentials, bundles);
						break;
					case 'replyInThread':
						result = await message.replyInThread.execute(this, i, authKind, credentials, bundles);
						break;
					default:
						throw new NodeOperationError(this.getNode(), `Unknown message operation: ${operation}`, { itemIndex: i });
				}
				break;
			}
			case 'card': {
				switch (operation) {
					case 'send':
						result = await card.send.execute(this, i, authKind, credentials, bundles);
						break;
					case 'update':
						result = await card.update.execute(this, i, authKind, credentials, bundles);
						break;
					default:
						throw new NodeOperationError(this.getNode(), `Unknown card operation: ${operation}`, { itemIndex: i });
				}
				break;
			}
			// invokeResponse is handled by the early return above.
			default: {
				throw new NodeOperationError(
					this.getNode(),
					`Unexpected resource "${String(resource)}" in the per-item loop. Invoke Response must be the only operation in this node execution; it responds to a single HTTP request and cannot be combined with other resources.`,
					{ itemIndex: i },
				);
			}
		}

		// Preserve full input item — binary, metadata, everything — and only
		// swap in the merged json. Earlier drafts dropped binary (manifest §12).
		out.push({ ...items[i], json: result, pairedItem: i });
	}

	return [out];
}
