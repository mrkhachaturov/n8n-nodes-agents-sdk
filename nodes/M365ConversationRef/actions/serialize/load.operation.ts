import type { IExecuteFunctions, IDataObject } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import {
	deserializeConversationRef,
	redactBearerInError,
	type SerializedConversationRef,
} from '../../../../shared/conversationRefSerializer';

/**
 * Serialize/load: reconstruct an envelope-shaped record from a previously
 * saved payload. Reads the payload from the input item's `payloadField`
 * (default "payload") and delegates the actual shape check + expiry derivation
 * to `deserializeConversationRef`. The `expired` flag is purely informational —
 * it is derived from `claimsExp < now`; downstream workflow logic decides how
 * to react.
 *
 * NOTE: the payload is expected to be an object. If a storage backend returns
 * it as a stringified JSON, the workflow must JSON.parse it before piping into
 * this node — `deserializeConversationRef` does not parse strings.
 */
export async function executeSerializeLoad(
	this: IExecuteFunctions,
	itemIndex: number,
): Promise<IDataObject[]> {
	const inputJson = this.getInputData()[itemIndex].json;
	const payloadField = this.getNodeParameter('payloadField', itemIndex, 'payload') as string;
	const rawPayload = (inputJson as Record<string, unknown>)[payloadField];

	if (rawPayload === undefined || rawPayload === null) {
		throw new NodeOperationError(
			this.getNode(),
			`Serialize/load: input item has no "${payloadField}" field. Populate "Payload Field" to point at where the saved payload lives.`,
			{ itemIndex },
		);
	}

	try {
		const result = deserializeConversationRef(rawPayload as SerializedConversationRef);
		return [
			{
				conversationReference: result.conversationReference,
				authContext: result.authContext,
				expired: result.expired,
			} as unknown as IDataObject,
		];
	} catch (err) {
		const redacted = redactBearerInError(err as Error);
		throw new NodeOperationError(this.getNode(), redacted.message, { itemIndex });
	}
}
