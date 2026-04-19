import type { IExecuteFunctions } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import type { Attachment } from '@microsoft/agents-activity';

function parseJsonParam(
	ctx: IExecuteFunctions,
	itemIndex: number,
	fieldName: string,
	raw: unknown,
): unknown {
	if (raw === undefined || raw === null) return {};
	if (typeof raw !== 'string') return raw;
	const trimmed = raw.trim();
	if (trimmed === '' || trimmed === '={}') return {};
	try {
		return JSON.parse(trimmed);
	} catch (err) {
		throw new NodeOperationError(
			ctx.getNode(),
			`Invalid JSON in ${fieldName}: ${(err as Error).message}`,
			{ itemIndex },
		);
	}
}

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
