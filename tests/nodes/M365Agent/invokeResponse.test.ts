import { describe, it, expect, vi } from 'vitest';
import * as invokeResponse from '../../../nodes/M365Agent/actions/invokeResponse';
import { respond } from '../../../nodes/M365Agent/actions/invokeResponse';

describe('invokeResponse resource description', () => {
	it('operation selector scoped to invokeResponse, alphabetical, noDataExpression', () => {
		const op = invokeResponse.description.find((p) => p.name === 'operation');
		expect(op?.noDataExpression).toBe(true);
		const shown = (op?.displayOptions?.show as Record<string, string[]> | undefined)?.resource;
		expect(shown).toEqual(['invokeResponse']);
		const values = (op?.options as { value: string }[] | undefined)?.map((o) => o.value);
		expect(values).toEqual(['respond']);
	});

	it('responseShape default is plain (Simple), noDataExpression', () => {
		const shape = invokeResponse.description.find((p) => p.name === 'responseShape');
		expect(shape?.default).toBe('plain');
		expect(shape?.noDataExpression).toBe(true);
	});

	it('cardResponseType shown only for adaptiveCard', () => {
		const card = invokeResponse.description.find((p) => p.name === 'cardResponseType');
		const shown = (card?.displayOptions?.show as Record<string, string[]> | undefined)
			?.responseShape;
		expect(shown).toEqual(['adaptiveCard']);
	});
});

describe('invokeResponse respond execute', () => {
	it('calls sendResponse exactly once regardless of item count', async () => {
		const sendResponse = vi.fn();
		const getNodeParameter = vi
			.fn()
			.mockImplementation((name: string, _: number, def?: unknown) => {
				if (name === 'responseShape') return 'plain';
				if (name === 'statusCode') return 200;
				if (name === 'body') return '{}';
				return def;
			});
		const ctx = {
			sendResponse,
			getNodeParameter,
			getInputData: () => [{ json: {} }, { json: {} }, { json: {} }],
			getNode: () => ({ name: 'test' }),
		} as unknown as import('n8n-workflow').IExecuteFunctions;

		await respond.execute(ctx);
		expect(sendResponse).toHaveBeenCalledTimes(1);
	});
});
