import { describe, it, expect, vi } from 'vitest';
import { NodeOperationError } from 'n8n-workflow';
import type { IExecuteFunctions } from 'n8n-workflow';
import { buildO365ConnectorAttachment } from '../../../../shared/card/buildAttachment/o365Connector';

function ctx(params: Record<string, unknown>): IExecuteFunctions {
	return {
		getNode: vi.fn().mockReturnValue({ name: 'T', id: 'id', type: 't', typeVersion: 1 }),
		getNodeParameter: vi.fn().mockImplementation((name: string, _i: unknown, def?: unknown) => {
			return name in params ? params[name] : def;
		}),
	} as unknown as IExecuteFunctions;
}

describe('buildO365ConnectorAttachment', () => {
	it('wraps the provided JSON as an o365 connector attachment', () => {
		const att = buildO365ConnectorAttachment(
			ctx({
				cardContent: '{"title":"Alert","summary":"S","sections":[]}',
			}),
			0,
		);
		expect(att.contentType).toBe('application/vnd.microsoft.teams.card.o365connector');
		expect((att.content as { title: string }).title).toBe('Alert');
	});

	it('accepts an already-parsed object', () => {
		const att = buildO365ConnectorAttachment(ctx({ cardContent: { title: 'R' } }), 0);
		expect((att.content as { title: string }).title).toBe('R');
	});

	it('throws NodeOperationError on invalid JSON', () => {
		expect(() => buildO365ConnectorAttachment(ctx({ cardContent: '{bad' }), 0)).toThrow(
			NodeOperationError,
		);
	});
});
