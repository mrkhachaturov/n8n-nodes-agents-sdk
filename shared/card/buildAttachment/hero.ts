import type { IExecuteFunctions } from 'n8n-workflow';
import type { Attachment } from '@microsoft/agents-activity';
import { CardFactory } from '@microsoft/agents-hosting';
import { buildButtons } from '../../cardBuilders/buttons';
import { buildImages } from '../../cardBuilders/images';
import { buildActions } from '../../cardBuilders/cardActionRow';

/**
 * Hero Card attachment builder. Mirrors `CardFactory.heroCard(title, text, images, buttons, other)`
 * from @microsoft/agents-hosting. Subtitle and Tap Action go into the `other` param.
 */
export function buildHeroAttachment(ctx: IExecuteFunctions, itemIndex: number): Attachment {
	const title = (ctx.getNodeParameter('title', itemIndex, '') as string) || '';
	const subtitle = (ctx.getNodeParameter('subtitle', itemIndex, '') as string) || undefined;
	const text = (ctx.getNodeParameter('text', itemIndex, '') as string) || undefined;
	const images = buildImages(ctx.getNodeParameter('images', itemIndex, {}) as never);
	const buttons = buildButtons(ctx.getNodeParameter('buttons', itemIndex, {}) as never);
	const options = ctx.getNodeParameter('options', itemIndex, {}) as {
		tapAction?: Parameters<typeof buildActions>[0] extends Array<infer R> | undefined ? R : never;
	};

	const tapRaw = options.tapAction;
	const tap =
		tapRaw && Object.keys(tapRaw as Record<string, unknown>).length > 0
			? buildActions([tapRaw])[0]
			: undefined;

	const other: Record<string, unknown> = {};
	if (subtitle) other.subtitle = subtitle;
	if (tap) other.tap = tap;

	return CardFactory.heroCard(
		title,
		text ?? '',
		images.length > 0 ? images : undefined,
		buttons.length > 0 ? buttons : undefined,
		other,
	);
}
