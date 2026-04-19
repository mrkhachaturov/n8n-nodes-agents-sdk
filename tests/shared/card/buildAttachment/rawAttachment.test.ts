import { describe, it, expect, vi } from 'vitest';
import { NodeOperationError } from 'n8n-workflow';
import type { IExecuteFunctions } from 'n8n-workflow';
import { buildRawAttachment } from '../../../../shared/card/buildAttachment/rawAttachment';

function ctx(params: Record<string, unknown>): IExecuteFunctions {
	return {
		getNode: vi.fn().mockReturnValue({ name: 'T', id: 'id', type: 't', typeVersion: 1 }),
		getNodeParameter: vi.fn().mockImplementation((name: string, _i: unknown, def?: unknown) => {
			return name in params ? params[name] : def;
		}),
	} as unknown as IExecuteFunctions;
}

describe('buildRawAttachment', () => {
	it('returns an Attachment with user-supplied contentType and content', () => {
		const att = buildRawAttachment(
			ctx({
				contentType: 'application/vnd.microsoft.card.adaptive',
				content: '{"type":"AdaptiveCard"}',
				options: {},
			}),
			0,
		);
		expect(att.contentType).toBe('application/vnd.microsoft.card.adaptive');
		expect(att.content).toEqual({ type: 'AdaptiveCard' });
	});

	it('threads through name/contentUrl/thumbnailUrl from options', () => {
		const att = buildRawAttachment(
			ctx({
				contentType: 'image/png',
				content: '={}',
				options: {
					name: 'file.png',
					contentUrl: 'https://cdn/file.png',
					thumbnailUrl: 'https://cdn/file-thumb.png',
				},
			}),
			0,
		);
		expect(att.name).toBe('file.png');
		expect(att.contentUrl).toBe('https://cdn/file.png');
		expect(att.thumbnailUrl).toBe('https://cdn/file-thumb.png');
	});

	it('throws NodeOperationError on empty Content Type', () => {
		expect(() => buildRawAttachment(ctx({ contentType: '', content: '{}' }), 0)).toThrow(
			/Content Type is required/,
		);
	});

	it('throws NodeOperationError on invalid JSON content', () => {
		expect(() => buildRawAttachment(ctx({ contentType: 'x', content: '{bad' }), 0)).toThrow(
			NodeOperationError,
		);
	});
});
