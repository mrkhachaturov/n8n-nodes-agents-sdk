import type { IExecuteFunctions } from 'n8n-workflow';
import type { Attachment } from '@microsoft/agents-activity';
import { parseJsonParam } from './parseJsonParam';

export function buildReceiptAttachment(ctx: IExecuteFunctions, itemIndex: number): Attachment {
	const cardContent = parseJsonParam(
		ctx,
		itemIndex,
		'Card Content',
		ctx.getNodeParameter('cardContent', itemIndex),
	) as Record<string, unknown>;
	return {
		contentType: 'application/vnd.microsoft.card.receipt',
		content: cardContent,
	};
}
