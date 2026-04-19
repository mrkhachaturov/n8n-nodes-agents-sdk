import { describe, it, expect, vi } from 'vitest';
import { NodeOperationError } from 'n8n-workflow';
import type { IExecuteFunctions } from 'n8n-workflow';
import { buildAnimationAttachment } from '../../../../shared/card/buildAttachment/animation';

function ctx(params: Record<string, unknown>): IExecuteFunctions {
	return {
		getNode: vi.fn().mockReturnValue({ name: 'T', id: 'id', type: 't', typeVersion: 1 }),
		getNodeParameter: vi.fn().mockImplementation((name: string, _i: unknown, def?: unknown) => {
			return name in params ? params[name] : def;
		}),
	} as unknown as IExecuteFunctions;
}

describe('buildAnimationAttachment', () => {
	it('throws NodeOperationError when no media URLs provided', () => {
		expect(() =>
			buildAnimationAttachment(ctx({ title: 'T', media: {}, buttons: {}, options: {} }), 0),
		).toThrow(NodeOperationError);
	});

	it('builds an animation card with media, poster image, and autostart from options', () => {
		const att = buildAnimationAttachment(
			ctx({
				title: 'Gif',
				subtitle: 'Loop',
				text: 'Body',
				posterImageUrl: 'https://poster',
				posterImageAltText: 'alt',
				media: { mediaItem: [{ url: 'https://a.gif' }] },
				buttons: {},
				options: {
					autoloop: true,
					autostart: true,
					shareable: false,
					aspect: '16:9',
					duration: 'PT30S',
					fallbackText: 'fallback',
					value: '={}',
				},
			}),
			0,
		);
		expect(att.contentType).toBe('application/vnd.microsoft.card.animation');
		const content = att.content as {
			title: string;
			image: { url: string; alt?: string };
			media: Array<{ url: string }>;
			autoloop: boolean;
			aspect: string;
		};
		expect(content.title).toBe('Gif');
		expect(content.image).toEqual({ url: 'https://poster', alt: 'alt' });
		expect(content.media).toEqual([{ url: 'https://a.gif' }]);
		expect(content.autoloop).toBe(true);
		expect(content.aspect).toBe('16:9');
	});

	it('parses Options.value JSON when non-empty', () => {
		const att = buildAnimationAttachment(
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
