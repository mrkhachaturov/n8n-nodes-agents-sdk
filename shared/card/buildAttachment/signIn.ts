import type { IExecuteFunctions } from 'n8n-workflow';
import type { Attachment } from '@microsoft/agents-activity';
import { CardFactory } from '@microsoft/agents-hosting';

export function buildSignInAttachment(ctx: IExecuteFunctions, itemIndex: number): Attachment {
	const title = ctx.getNodeParameter('title', itemIndex, '') as string;
	const url = ctx.getNodeParameter('url', itemIndex, '') as string;
	const text = (ctx.getNodeParameter('text', itemIndex, '') as string) || undefined;
	return CardFactory.signinCard(title, url, text);
}
