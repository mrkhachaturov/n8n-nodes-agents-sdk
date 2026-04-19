import type { IExecuteFunctions } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import type { Attachment } from '@microsoft/agents-activity';

function parseJsonParam(
	ctx: IExecuteFunctions,
	itemIndex: number,
	fieldName: string,
	raw: unknown,
): Record<string, unknown> {
	if (raw === undefined || raw === null) return {};
	if (typeof raw !== 'string') return raw as Record<string, unknown>;
	try {
		return JSON.parse(raw) as Record<string, unknown>;
	} catch (err) {
		throw new NodeOperationError(
			ctx.getNode(),
			`Invalid JSON in ${fieldName}: ${(err as Error).message}`,
			{ itemIndex },
		);
	}
}

export function buildReceiptAttachment(ctx: IExecuteFunctions, itemIndex: number): Attachment {
	const cardContent = parseJsonParam(
		ctx,
		itemIndex,
		'Card Content',
		ctx.getNodeParameter('cardContent', itemIndex),
	);
	return {
		contentType: 'application/vnd.microsoft.card.receipt',
		content: cardContent,
	};
}
