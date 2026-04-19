import { describe, it, expect, vi } from 'vitest';
import { NodeOperationError } from 'n8n-workflow';
import type { IExecuteFunctions } from 'n8n-workflow';
import { buildAudioAttachment } from '../../../../shared/card/buildAttachment/audio';

function ctx(params: Record<string, unknown>): IExecuteFunctions {
	return {
		getNode: vi.fn().mockReturnValue({ name: 'T', id: 'id', type: 't', typeVersion: 1 }),
		getNodeParameter: vi.fn().mockImplementation((name: string, _i: unknown, def?: unknown) => {
			return name in params ? params[name] : def;
		}),
	} as unknown as IExecuteFunctions;
}

describe('buildAudioAttachment', () => {
	it('throws NodeOperationError when no media URLs provided', () => {
		expect(() =>
			buildAudioAttachment(ctx({ title: 'T', media: {}, buttons: {}, options: {} }), 0),
		).toThrow(NodeOperationError);
	});

	it('builds an audio card with media, poster image, and autostart from options', () => {
		const att = buildAudioAttachment(
			ctx({
				title: 'Track',
				subtitle: 'Loop',
				text: 'Body',
				posterImageUrl: 'https://poster',
				posterImageAltText: 'alt',
				media: { mediaItem: [{ url: 'https://a.mp3' }] },
				buttons: {},
				options: {
					autoloop: true,
					autostart: true,
					shareable: false,
					aspect: '16:9',
					duration: 'PT2M30S',
					fallbackText: 'fallback',
					value: '={}',
				},
			}),
			0,
		);
		expect(att.contentType).toBe('application/vnd.microsoft.card.audio');
		const content = att.content as {
			title: string;
			image: { url: string; alt?: string };
			media: Array<{ url: string }>;
			autoloop: boolean;
			aspect: string;
		};
		expect(content.title).toBe('Track');
		expect(content.image).toEqual({ url: 'https://poster', alt: 'alt' });
		expect(content.media).toEqual([{ url: 'https://a.mp3' }]);
		expect(content.autoloop).toBe(true);
		expect(content.aspect).toBe('16:9');
	});

	it('parses Options.value JSON when non-empty', () => {
		const att = buildAudioAttachment(
			ctx({
				title: 'T',
				media: { mediaItem: [{ url: 'https://m' }] },
				options: { value: '{"ref":1}' },
			}),
			0,
		);
		expect((att.content as { value: unknown }).value).toEqual({ ref: 1 });
	});
});
