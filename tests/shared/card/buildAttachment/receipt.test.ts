import { describe, it, expect, vi } from 'vitest';
import { NodeOperationError } from 'n8n-workflow';
import type { IExecuteFunctions } from 'n8n-workflow';
import { buildReceiptAttachment } from '../../../../shared/card/buildAttachment/receipt';

function ctx(params: Record<string, unknown>): IExecuteFunctions {
	return {
		getNode: vi.fn().mockReturnValue({ name: 'T', id: 'id', type: 't', typeVersion: 1 }),
		getNodeParameter: vi.fn().mockImplementation((name: string, _i: unknown, def?: unknown) => {
			return name in params ? params[name] : def;
		}),
	} as unknown as IExecuteFunctions;
}

describe('buildReceiptAttachment', () => {
	it('wraps the provided JSON as a receipt attachment', () => {
		const att = buildReceiptAttachment(
			ctx({
				cardContent: '{"title":"Receipt #1","items":[{"title":"X","price":"10"}],"total":"10"}',
			}),
			0,
		);
		expect(att.contentType).toBe('application/vnd.microsoft.card.receipt');
		expect((att.content as { title: string }).title).toBe('Receipt #1');
	});

	it('accepts an already-parsed object', () => {
		const att = buildReceiptAttachment(ctx({ cardContent: { title: 'R' } }), 0);
		expect((att.content as { title: string }).title).toBe('R');
	});

	it('throws NodeOperationError on invalid JSON', () => {
		expect(() => buildReceiptAttachment(ctx({ cardContent: '{bad' }), 0)).toThrow(
			NodeOperationError,
		);
	});
});
