import { describe, it, expect, vi } from 'vitest';
import type { IExecuteFunctions } from 'n8n-workflow';
import { buildHeroAttachment } from '../../../../shared/card/buildAttachment/hero';

function ctx(params: Record<string, unknown>): IExecuteFunctions {
	return {
		getNode: vi.fn().mockReturnValue({ name: 'T', id: 'id', type: 't', typeVersion: 1 }),
		getNodeParameter: vi.fn().mockImplementation((name: string, _i: unknown, def?: unknown) => {
			return name in params ? params[name] : def;
		}),
	} as unknown as IExecuteFunctions;
}

describe('buildHeroAttachment', () => {
	it('builds a hero card with title, subtitle, text, images, buttons, and tap', () => {
		const att = buildHeroAttachment(
			ctx({
				title: 'Main',
				subtitle: 'Sub',
				text: 'Body',
				images: {
					image: [{ url: 'https://img', alt: 'alt' }],
				},
				buttons: {
					button: [{ type: 'openUrl', title: 'Open', value: 'https://x' }],
				},
				options: {
					fallbackText: 'fallback',
					tapAction: { type: 'openUrl', title: 'Main Open', value: 'https://main' },
				},
			}),
			0,
		);
		expect(att.contentType).toBe('application/vnd.microsoft.card.hero');
		const content = att.content as {
			title: string;
			subtitle: string;
			text: string;
			images: Array<{ url: string; alt?: string }>;
			buttons: Array<{ type: string }>;
			tap: { type: string; value: string };
		};
		expect(content.title).toBe('Main');
		expect(content.subtitle).toBe('Sub');
		expect(content.text).toBe('Body');
		expect(content.images).toEqual([{ url: 'https://img', alt: 'alt' }]);
		expect(content.buttons).toHaveLength(1);
		expect(content.tap).toEqual({ type: 'openUrl', title: 'Main Open', value: 'https://main' });
	});

	it('omits empty image/button arrays and omits tap when absent', () => {
		const att = buildHeroAttachment(ctx({ title: 'T', options: {} }), 0);
		const content = att.content as Record<string, unknown>;
		expect(content.title).toBe('T');
		expect(content.images).toBeUndefined();
		expect(content.buttons).toBeUndefined();
		expect(content.tap).toBeUndefined();
	});
});
