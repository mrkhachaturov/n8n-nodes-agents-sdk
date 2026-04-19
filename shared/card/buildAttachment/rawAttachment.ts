import type { IExecuteFunctions } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import type { Attachment } from '@microsoft/agents-activity';
import { parseJsonParam } from './parseJsonParam';

export function buildRawAttachment(ctx: IExecuteFunctions, itemIndex: number): Attachment {
	const contentType = (ctx.getNodeParameter('contentType', itemIndex, '') as string).trim();
	if (!contentType) {
		throw new NodeOperationError(ctx.getNode(), 'Content Type is required', { itemIndex });
	}
	const content = parseJsonParam(
		ctx,
		itemIndex,
		'Content',
		ctx.getNodeParameter('content', itemIndex, '={}'),
	);
	const options = ctx.getNodeParameter('options', itemIndex, {}) as {
		name?: string;
		contentUrl?: string;
		thumbnailUrl?: string;
	};
	const attachment: Attachment = { contentType, content };
	if (options.name) attachment.name = options.name;
	if (options.contentUrl) attachment.contentUrl = options.contentUrl;
	if (options.thumbnailUrl) attachment.thumbnailUrl = options.thumbnailUrl;
	return attachment;
}
