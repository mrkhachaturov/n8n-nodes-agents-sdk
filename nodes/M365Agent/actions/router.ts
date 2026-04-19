import type { IDataObject, IExecuteFunctions, INodeExecutionData } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import type { BotConnectorBundle } from '../../../shared/botConnector';
import type { AuthKind, M365ClassicBotCred, M365Agent365Cred } from '../../../shared/types';

import * as message from './message';
import * as cardAdaptive from './card-adaptive';
import * as cardAnimation from './card-animation';
import * as cardAudio from './card-audio';
import * as cardHero from './card-hero';
import * as cardSignIn from './card-sign-in';
import * as cardThumbnail from './card-thumbnail';
import * as cardVideo from './card-video';
import * as invokeResponse from './invokeResponse';
import { makeBundleKey } from './bundleKey';

// Re-export so external callers (tests, docs) can still import from 'router'.
export { makeBundleKey };

type Resource =
	| 'message'
	| 'adaptiveCard'
	| 'animationCard'
	| 'audioCard'
	| 'heroCard'
	| 'signInCard'
	| 'thumbnailCard'
	| 'videoCard'
	| 'invokeResponse';

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
		for (let i = 0; i < items.length; i++) {
			const r = this.getNodeParameter('resource', i) as Resource;
			const op = this.getNodeParameter('operation', i) as string;
			if (r !== 'invokeResponse') {
				throw new NodeOperationError(
					this.getNode(),
					`Invoke Response must be the only resource for the whole batch. Item ${i} resolves to resource "${String(r)}".`,
					{ itemIndex: i },
				);
			}
			if (op !== 'respond') {
				throw new NodeOperationError(
					this.getNode(),
					`Unknown invokeResponse operation: ${op}`,
					{ itemIndex: i },
				);
			}
		}
		await invokeResponse.respond.execute(this);
		return [items.map((item, i) => ({ ...item, pairedItem: { item: i } }))];
	}

	// Credentials are only needed for message/card operations (both talk to the
	// Bot Connector via MSAL). Fetching after the invokeResponse early-return
	// avoids a pointless vault roundtrip on the response-only path.
	// authKind drives which credential family to fetch — classicBot uses the
	// legacy M365AgentApi credential (App ID + secret/MSI), agent365 uses the
	// new M365Agent365Api credential (Blueprint App + sidecar/inline transport).
	const authKind = this.getNodeParameter('authKind', 0, 'classicBot') as AuthKind;
	const credentialsRaw =
		authKind === 'classicBot'
			? await this.getCredentials('m365AgentApi')
			: await this.getCredentials('m365Agent365Api');
	const credentials = credentialsRaw as unknown as M365ClassicBotCred | M365Agent365Cred;

	const bundles = new Map<string, BotConnectorBundle>();

	for (let i = 0; i < items.length; i++) {
		const resource = this.getNodeParameter('resource', i) as Resource;
		const operation = this.getNodeParameter('operation', i) as string;

		try {
			let result: IDataObject;
			switch (resource) {
				case 'message': {
					switch (operation) {
						case 'send':
							result = await message.send.execute.call(this, i, authKind, credentials, bundles);
							break;
						case 'reply':
							result = await message.reply.execute.call(this, i, authKind, credentials, bundles);
							break;
						case 'update':
							result = await message.update.execute.call(this, i, authKind, credentials, bundles);
							break;
						case 'delete':
							result = await message.deleteMessage.execute.call(
								this,
								i,
								authKind,
								credentials,
								bundles,
							);
							break;
						case 'replyInThread':
							result = await message.replyInThread.execute.call(
								this,
								i,
								authKind,
								credentials,
								bundles,
							);
							break;
						default:
							throw new NodeOperationError(
								this.getNode(),
								`Unknown message operation: ${operation}`,
								{ itemIndex: i },
							);
					}
					break;
				}
				case 'adaptiveCard': {
					switch (operation) {
						case 'send':
							result = await cardAdaptive.send.execute.call(
								this,
								i,
								authKind,
								credentials,
								bundles,
							);
							break;
						case 'update':
							result = await cardAdaptive.update.execute.call(
								this,
								i,
								authKind,
								credentials,
								bundles,
							);
							break;
						default:
							throw new NodeOperationError(
								this.getNode(),
								`Unknown adaptiveCard operation: ${operation}`,
								{ itemIndex: i },
							);
					}
					break;
				}
				case 'animationCard': {
					switch (operation) {
						case 'send':
							result = await cardAnimation.send.execute.call(
								this,
								i,
								authKind,
								credentials,
								bundles,
							);
							break;
						case 'update':
							result = await cardAnimation.update.execute.call(
								this,
								i,
								authKind,
								credentials,
								bundles,
							);
							break;
						default:
							throw new NodeOperationError(
								this.getNode(),
								`Unknown animationCard operation: ${operation}`,
								{ itemIndex: i },
							);
					}
					break;
				}
				case 'audioCard': {
					switch (operation) {
						case 'send':
							result = await cardAudio.send.execute.call(
								this,
								i,
								authKind,
								credentials,
								bundles,
							);
							break;
						case 'update':
							result = await cardAudio.update.execute.call(
								this,
								i,
								authKind,
								credentials,
								bundles,
							);
							break;
						default:
							throw new NodeOperationError(
								this.getNode(),
								`Unknown audioCard operation: ${operation}`,
								{ itemIndex: i },
							);
					}
					break;
				}
				case 'heroCard': {
					switch (operation) {
						case 'send':
							result = await cardHero.send.execute.call(this, i, authKind, credentials, bundles);
							break;
						case 'update':
							result = await cardHero.update.execute.call(
								this,
								i,
								authKind,
								credentials,
								bundles,
							);
							break;
						default:
							throw new NodeOperationError(
								this.getNode(),
								`Unknown heroCard operation: ${operation}`,
								{ itemIndex: i },
							);
					}
					break;
				}
				case 'signInCard': {
					switch (operation) {
						case 'send':
							result = await cardSignIn.send.execute.call(
								this,
								i,
								authKind,
								credentials,
								bundles,
							);
							break;
						case 'update':
							result = await cardSignIn.update.execute.call(
								this,
								i,
								authKind,
								credentials,
								bundles,
							);
							break;
						default:
							throw new NodeOperationError(
								this.getNode(),
								`Unknown signInCard operation: ${operation}`,
								{ itemIndex: i },
							);
					}
					break;
				}
				case 'thumbnailCard': {
					switch (operation) {
						case 'send':
							result = await cardThumbnail.send.execute.call(
								this,
								i,
								authKind,
								credentials,
								bundles,
							);
							break;
						case 'update':
							result = await cardThumbnail.update.execute.call(
								this,
								i,
								authKind,
								credentials,
								bundles,
							);
							break;
						default:
							throw new NodeOperationError(
								this.getNode(),
								`Unknown thumbnailCard operation: ${operation}`,
								{ itemIndex: i },
							);
					}
					break;
				}
				case 'videoCard': {
					switch (operation) {
						case 'send':
							result = await cardVideo.send.execute.call(
								this,
								i,
								authKind,
								credentials,
								bundles,
							);
							break;
						case 'update':
							result = await cardVideo.update.execute.call(
								this,
								i,
								authKind,
								credentials,
								bundles,
							);
							break;
						default:
							throw new NodeOperationError(
								this.getNode(),
								`Unknown videoCard operation: ${operation}`,
								{ itemIndex: i },
							);
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
			// swap in the merged json. Matches manifest §12 "never drop binary".
			out.push({ ...items[i], json: result, pairedItem: { item: i } });
		} catch (err) {
			if (this.continueOnFail()) {
				// Full input item preserved on the error path too — only json is merged.
				// Inlined (err as Error).message to avoid shadowing the `message` module import.
				out.push({
					...items[i],
					json: { ...(items[i].json as IDataObject), error: (err as Error).message },
					pairedItem: { item: i },
				});
				continue;
			}
			throw err;
		}
	}

	return [out];
}
