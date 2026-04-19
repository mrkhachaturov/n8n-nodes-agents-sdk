import { describe, it, expect, vi } from 'vitest';
import { NodeOperationError } from 'n8n-workflow';
import type { IExecuteFunctions } from 'n8n-workflow';
import { buildAdaptiveAttachment } from '../../../../shared/card/buildAttachment/adaptive';

function ctx(params: Record<string, unknown>): IExecuteFunctions {
	return {
		getNode: vi.fn().mockReturnValue({ name: 'T', id: 'id', type: 't', typeVersion: 1 }),
		getNodeParameter: vi.fn().mockImplementation((name: string, _i: unknown, def?: unknown) => {
			if (name in params) return params[name];
			return def;
		}),
	} as unknown as IExecuteFunctions;
}

describe('buildAdaptiveAttachment', () => {
	it('expands the card template with bindingData and returns an Adaptive attachment', () => {
		const attachment = buildAdaptiveAttachment(
			ctx({
				cardTemplate:
					'{"type":"AdaptiveCard","version":"1.4","body":[{"type":"TextBlock","text":"${title}"}]}',
				bindingData: { title: 'Hello' },
			}),
			0,
		);
		expect(attachment.contentType).toBe('application/vnd.microsoft.card.adaptive');
		expect(JSON.stringify(attachment.content)).toContain('Hello');
	});

	it('accepts pre-parsed object templates', () => {
		const attachment = buildAdaptiveAttachment(
			ctx({
				cardTemplate: { type: 'AdaptiveCard', version: '1.4', body: [] },
				bindingData: {},
			}),
			0,
		);
		expect(attachment.content).toMatchObject({ type: 'AdaptiveCard', version: '1.4' });
	});

	it('defaults bindingData to {} when not provided', () => {
		expect(() =>
			buildAdaptiveAttachment(
				ctx({ cardTemplate: '{"type":"AdaptiveCard","version":"1.4","body":[]}' }),
				0,
			),
		).not.toThrow();
	});

	it('throws NodeOperationError on invalid JSON in cardTemplate', () => {
		expect(() =>
			buildAdaptiveAttachment(ctx({ cardTemplate: '{not-json', bindingData: {} }), 0),
		).toThrow(NodeOperationError);
	});

	it('throws NodeOperationError on template expansion failure', async () => {
		// Test the expansion branch deterministically by mocking Template.expand
		// to throw — avoids depending on adaptivecards-templating parsing quirks
		// for any specific malformed expression. Uses a pre-parsed object as
		// cardTemplate so the JSON.parse branch is skipped and we hit Template.expand.
		const { Template } = await import('adaptivecards-templating');
		const spy = vi.spyOn(Template.prototype, 'expand').mockImplementation(() => {
			throw new Error('simulated template parsing failed');
		});
		try {
			expect(() =>
				buildAdaptiveAttachment(
					ctx({
						cardTemplate: { type: 'AdaptiveCard', version: '1.4', body: [] },
						bindingData: {},
					}),
					0,
				),
			).toThrow(/Card Template expansion failed/);
		} finally {
			spy.mockRestore();
		}
	});
});
