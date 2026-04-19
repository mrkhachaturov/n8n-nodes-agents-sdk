import type { IDataObject, IExecuteFunctions, JsonObject } from 'n8n-workflow';
import { NodeApiError, NodeOperationError } from 'n8n-workflow';
import type { Activity, Attachment } from '@microsoft/agents-activity';

import { acquireOutboundToken } from '../auth/router';
import { CardBuildError } from '../cardBuilders/cardActionRow';
import { createConnectorFromBearer, type BotConnectorBundle } from '../botConnector';
import type {
  AuthKind,
  IdentityMode,
  M365ClassicBotCred,
  M365Agent365Cred,
} from '../types';
import { resolveConversationReference } from '../../nodes/M365Agent/descriptions/conversationReference';
import { makeBundleKey } from '../../nodes/M365Agent/actions/bundleKey';

export interface DispatchSendArgs {
  ctx: IExecuteFunctions;
  itemIndex: number;
  /** Internal resource key for error messages — e.g. "heroCard", "adaptiveCard". */
  resourceLabel: string;
  authKind: AuthKind;
  credentials: M365ClassicBotCred | M365Agent365Cred;
  bundles: Map<string, BotConnectorBundle>;
  /** Per-resource form → Attachment normalizer. */
  buildAttachment: (ctx: IExecuteFunctions, itemIndex: number) => Attachment;
}

export async function dispatchSend(args: DispatchSendArgs): Promise<IDataObject> {
  const { ctx, itemIndex, resourceLabel, authKind, credentials, bundles, buildAttachment } = args;

  const item = ctx.getInputData()[itemIndex].json as IDataObject;
  const ref = resolveConversationReference(ctx, itemIndex, item);

  // Build the attachment BEFORE acquiring a token — validation errors short-circuit.
  // Apply the spec §5.2 error-class rule:
  //   * NodeOperationError passes through unchanged (user config — yellow banner).
  //   * CardBuildError (from shared JSON parsers in cardActionRow) is normalized
  //     to NodeOperationError — it always represents invalid user JSON, not API failure.
  //   * Any other throwable (unexpected programmer crash, template-library exception
  //     that escaped the per-builder catch, etc.) becomes NodeApiError labeled as a
  //     card build failure so the caller sees a typed error class.
  let attachment: Attachment;
  try {
    attachment = buildAttachment(ctx, itemIndex);
  } catch (err) {
    if (err instanceof NodeOperationError) throw err;
    if (err instanceof CardBuildError) {
      throw new NodeOperationError(ctx.getNode(), (err as Error).message, { itemIndex });
    }
    const e = err as Error;
    throw new NodeApiError(ctx.getNode(), { message: e.message } as JsonObject, {
      message: `M365 Agent ${resourceLabel} — card build failed: ${e.message}`,
    });
  }

  const options = ctx.getNodeParameter('options', itemIndex, {}) as IDataObject;
  const activity: Partial<Activity> = {
    type: 'message',
    attachments: [attachment],
    ...(options.fallbackText ? { text: options.fallbackText as string } : {}),
  };

  // Auth routing (mirrors card/send.operation.ts lines 96-124).
  const identityMode: IdentityMode =
    authKind === 'agent365'
      ? (ctx.getNodeParameter('identityMode', itemIndex, 'autonomous') as IdentityMode)
      : 'autonomous';

  let agentUsername: string | undefined;
  let agentUserId: string | undefined;
  if (authKind === 'agent365' && identityMode === 'agentUser') {
    const mode = ctx.getNodeParameter('userSelectorMode', itemIndex, 'byUpn') as string;
    if (mode === 'byUpn') {
      agentUsername =
        (ctx.getNodeParameter('agentUsername', itemIndex, '') as string) || undefined;
    } else {
      agentUserId = (ctx.getNodeParameter('agentUserId', itemIndex, '') as string) || undefined;
    }
  }

  const downstreamApi: string =
    authKind === 'agent365'
      ? ((credentials as M365Agent365Cred).outboundDownstreamApi ?? 'MessagingBotApi')
      : 'BotFramework';

  const bundleKey = makeBundleKey(ref.serviceUrl, authKind, identityMode, agentUsername, agentUserId);
  let bundle = bundles.get(bundleKey);
  if (!bundle) {
    // Let NodeOperationError from backends (agent365Inline agentUser / interactiveOBO
    // rejection) propagate unchanged. Wrap only non-NodeOperationError failures.
    try {
      const { authorizationHeader } = await acquireOutboundToken({
        authKind,
        credentials,
        identityMode,
        downstreamApi,
        agentUsername,
        agentUserId,
      });
      bundle = createConnectorFromBearer(ref.serviceUrl, authorizationHeader);
      bundles.set(bundleKey, bundle);
    } catch (err) {
      if (err instanceof NodeOperationError) throw err;
      const e = err as Error;
      throw new NodeApiError(ctx.getNode(), { message: e.message } as JsonObject, {
        message: `M365 Agent token acquire failed: ${e.message}`,
      });
    }
  }

  try {
    const result = await bundle.client.sendToConversation(
      ref.conversation.id,
      activity as Activity,
    );
    return { ...item, cardSendResult: { id: result.id } };
  } catch (err) {
    // Defensive: connector today throws Axios errors, not NodeOperationError.
    // The guard is here so future connector wrappers (e.g. agent365 sidecar) that
    // pre-classify auth-layer user errors as NodeOperationError propagate cleanly.
    if (err instanceof NodeOperationError) throw err;
    const e = err as Error;
    throw new NodeApiError(ctx.getNode(), { message: e.message } as JsonObject, {
      message: `M365 Agent ${resourceLabel} send failed: ${e.message}`,
    });
  }
}
