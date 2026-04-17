import type { IExecuteFunctions, INodeProperties } from 'n8n-workflow';

export const description: INodeProperties[] = [];

/**
 * Invoke Response writes ONE HTTP response for the whole batch (manifest §9 —
 * mirrors RespondToWebhook). Unlike the message/card operations, this one does
 * NOT take itemIndex — parameters are read from item 0 and the response goes
 * back through ctx.sendResponse() exactly once.
 */
export async function execute(ctx: IExecuteFunctions): Promise<void> {
	const responseShape = ctx.getNodeParameter('responseShape', 0) as 'plain' | 'adaptiveCard';
	const statusCode = (ctx.getNodeParameter('statusCode', 0) as number) ?? 200;
	const bodyRaw = ctx.getNodeParameter('body', 0, '{}') as unknown;
	const body: Record<string, unknown> =
		typeof bodyRaw === 'string'
			? JSON.parse(bodyRaw || '{}')
			: (bodyRaw as Record<string, unknown>);

	let responsePayload: Record<string, unknown>;
	if (responseShape === 'adaptiveCard') {
		const cardResponseType = ctx.getNodeParameter('cardResponseType', 0) as string;
		responsePayload = {
			statusCode,
			type: cardResponseType,
			value: body,
		};
	} else {
		responsePayload = {
			status: statusCode,
			body,
		};
	}

	ctx.sendResponse({
		body: responsePayload,
		statusCode: 200, // the webhook envelope always returns 200 to Azure Bot Service; the semantic status lives inside the body
		headers: { 'content-type': 'application/json' },
	});
}
