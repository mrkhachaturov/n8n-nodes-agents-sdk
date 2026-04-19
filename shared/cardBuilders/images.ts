import type { INodeProperties } from 'n8n-workflow';
import type { CardAction } from '@microsoft/agents-activity';
import { cardActionFields, buildActions } from './cardActionRow';

interface CardImage {
	url: string;
	alt?: string;
	tap?: CardAction;
}

/**
 * Images builder — fixedCollection(multipleValues). Each row has url + alt +
 * Options.tapAction (a single CardAction). Tap is a singleton per SDK CardImage.
 * Used by Hero and Thumbnail resources.
 */
export function imagesField(): INodeProperties {
	return {
		displayName: 'Images',
		name: 'images',
		type: 'fixedCollection',
		typeOptions: { multipleValues: true },
		default: {},
		options: [
			{
				displayName: 'Image',
				name: 'image',
				values: [
					{
						displayName: 'URL',
						name: 'url',
						type: 'string',
						required: true,
						default: '',
						description: 'Image URL',
					},
					{
						displayName: 'Alt Text',
						name: 'alt',
						type: 'string',
						default: '',
						description: 'Accessibility alt text',
					},
					{
						displayName: 'Options',
						name: 'options',
						type: 'collection',
						placeholder: 'Add option',
						default: {},
						options: [
							{
								displayName: 'Tap Action',
								name: 'tapAction',
								type: 'collection',
								placeholder: 'Add Tap Action',
								default: {},
								description: 'Action invoked when the user taps the image',
								options: cardActionFields(),
							},
						],
					},
				],
			},
		],
	};
}

interface ImageRow {
	url?: string;
	alt?: string;
	options?: {
		tapAction?: Parameters<typeof buildActions>[0] extends Array<infer R> | undefined ? R : never;
	};
}

interface ImagesCollection {
	image?: ImageRow[];
}

/** Convert the raw Images fixedCollection to `CardImage[]`. */
export function buildImages(raw: ImagesCollection | undefined): CardImage[] {
	const rows = raw?.image;
	if (!rows || rows.length === 0) return [];
	return rows.map((row) => {
		const image: CardImage = { url: row.url ?? '' };
		if (row.alt) image.alt = row.alt;
		const tapRaw = row.options?.tapAction;
		if (tapRaw && Object.keys(tapRaw as Record<string, unknown>).length > 0) {
			const [tap] = buildActions([tapRaw]);
			if (tap) image.tap = tap;
		}
		return image;
	});
}
