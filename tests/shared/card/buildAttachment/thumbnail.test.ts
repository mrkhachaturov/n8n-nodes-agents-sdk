import { describe, it, expect, vi } from 'vitest';
import type { IExecuteFunctions } from 'n8n-workflow';
import { buildThumbnailAttachment } from '../../../../shared/card/buildAttachment/thumbnail';

function ctx(params: Record<string, unknown>): IExecuteFunctions {
	return {
		getNode: vi.fn().mockReturnValue({ name: 'T', id: 'id', type: 't', typeVersion: 1 }),
		getNodeParameter: vi.fn().mockImplementation((name: string, _i: unknown, def?: unknown) => {
			return name in params ? params[name] : def;
		}),
	} as unknown as IExecuteFunctions;
}

describe('buildThumbnailAttachment', () => {
	it('produces a thumbnail contentType attachment', () => {
		const att = buildThumbnailAttachment(
			ctx({
				title: 'Thumb',
				subtitle: 'Sub',
				text: '',
				images: {},
				buttons: {},
				options: {},
			}),
			0,
		);
		expect(att.contentType).toBe('application/vnd.microsoft.card.thumbnail');
		expect((att.content as { title: string }).title).toBe('Thumb');
	});

	it('reads tap from the fixedCollection-wrapped Options.tapAction.action shape', () => {
		const att = buildThumbnailAttachment(
			ctx({
				title: 'Thumb',
				options: {
					tapAction: {
						action: { type: 'openUrl', title: 'Tap', value: 'https://x' },
					},
				},
			}),
			0,
		);
		const content = att.content as { tap?: { type: string; value: string } };
		expect(content.tap).toEqual({ type: 'openUrl', title: 'Tap', value: 'https://x' });
	});
});
