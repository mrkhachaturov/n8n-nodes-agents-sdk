import type { IExecuteFunctions } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import type { Attachment } from '@microsoft/agents-activity';
import { CardFactory } from '@microsoft/agents-hosting';
import { buildButtons } from '../../cardBuilders/buttons';
import { buildMedia } from '../../cardBuilders/media';

interface MediaCardOptions {
	autoloop?: boolean;
	autostart?: boolean;
	shareable?: boolean;
	aspect?: string;
	duration?: string;
	fallbackText?: string;
	value?: string | Record<string, unknown>;
}

function parseOptionalJson(raw: unknown): unknown {
	if (raw === undefined || raw === null) return undefined;
	if (typeof raw === 'object') return raw;
	if (typeof raw !== 'string') return undefined;
	const trimmed = raw.trim();
	if (trimmed === '' || trimmed === '={}' || trimmed === '{}') return undefined;
	return JSON.parse(trimmed);
}

export function buildAudioAttachment(ctx: IExecuteFunctions, itemIndex: number): Attachment {
	const title = (ctx.getNodeParameter('title', itemIndex, '') as string) || '';
	const subtitle = (ctx.getNodeParameter('subtitle', itemIndex, '') as string) || undefined;
	const text = (ctx.getNodeParameter('text', itemIndex, '') as string) || undefined;
	const posterUrl = (ctx.getNodeParameter('posterImageUrl', itemIndex, '') as string) || undefined;
	const posterAlt =
		(ctx.getNodeParameter('posterImageAltText', itemIndex, '') as string) || undefined;
	const media = buildMedia(ctx.getNodeParameter('media', itemIndex, {}) as never);
	if (media.length === 0) {
		throw new NodeOperationError(ctx.getNode(), 'Audio Card requires at least one media URL', {
			itemIndex,
		});
	}
	const buttons = buildButtons(ctx.getNodeParameter('buttons', itemIndex, {}) as never);
	const options = ctx.getNodeParameter('options', itemIndex, {}) as MediaCardOptions;

	const other: Record<string, unknown> = {};
	if (subtitle) other.subtitle = subtitle;
	if (text) other.text = text;
	if (posterUrl) other.image = posterAlt ? { url: posterUrl, alt: posterAlt } : { url: posterUrl };
	if (options.autoloop !== undefined) other.autoloop = options.autoloop;
	if (options.autostart !== undefined) other.autostart = options.autostart;
	if (options.shareable !== undefined) other.shareable = options.shareable;
	if (options.aspect) other.aspect = options.aspect;
	if (options.duration) other.duration = options.duration;
	let value: unknown;
	try {
		value = parseOptionalJson(options.value);
	} catch (err) {
		throw new NodeOperationError(
			ctx.getNode(),
			`Invalid JSON in Value: ${(err as Error).message}`,
			{ itemIndex },
		);
	}
	if (value !== undefined) other.value = value;

	return CardFactory.audioCard(title, media, buttons.length > 0 ? buttons : undefined, other);
}
