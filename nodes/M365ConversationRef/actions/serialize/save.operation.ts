import type { IExecuteFunctions, IDataObject } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import {
	serializeConversationRef,
	redactBearerInError,
} from '../../../../shared/conversationRefSerializer';
import type { ItemEnvelope } from '../../../../shared/types';

/**
 * Serialize/save: flatten the inbound ItemEnvelope into the narrow persistence
 * payload defined by `serializeConversationRef` (T3-conversation-reference-persistence).
 * Bearer is excluded unless the `includeBearer` toggle is on. Any error thrown
 * by the helper is passed through `redactBearerInError` before being wrapped in
 * NodeOperationError so execution logs never leak token material.
 */
export async function executeSerializeSave(
	this: IExecuteFunctions,
	itemIndex: number,
): Promise<IDataObject[]> {
	const envelope = this.getInputData()[itemIndex].json as unknown as ItemEnvelope;
	const includeBearer = this.getNodeParameter('includeBearer', itemIndex, false) as boolean;
	try {
		const payload = serializeConversationRef(envelope, { includeBearer });
		return [payload as unknown as IDataObject];
	} catch (err) {
		const redacted = redactBearerInError(err as Error);
		throw new NodeOperationError(this.getNode(), redacted.message, { itemIndex });
	}
}
