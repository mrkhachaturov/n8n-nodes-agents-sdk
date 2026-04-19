import { describe, it, expect, vi } from 'vitest';
import type { IExecuteFunctions } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import * as invokeResponse from '../../../nodes/M365Agent/actions/invokeResponse';
import { respond } from '../../../nodes/M365Agent/actions/invokeResponse';
import { M365Agent } from '../../../nodes/M365Agent/M365Agent.node';
import { makeExecuteContext, makeCredentials } from '../../helpers/makeContext';

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

describe('invokeResponse router output — pairedItem object form + mixed-batch rejection', () => {
	it('router early-return emits pairedItem: { item: i } for each input item', async () => {
		const node = new M365Agent();
		const sendResponse = vi.fn();
		const ctx = makeExecuteContext({
			inputItems: [{ key: 'a' }, { key: 'b' }],
			credentials: makeCredentials(),
			parameters: {
				resource: 'invokeResponse',
				operation: 'respond',
				responseShape: 'plain',
				statusCode: 200,
				body: '{}',
			},
		});
		(ctx as unknown as { sendResponse: typeof sendResponse }).sendResponse = sendResponse;
		const result = await node.execute.call(ctx as unknown as IExecuteFunctions);
		expect(result[0]).toHaveLength(2);
		expect(result[0][0].pairedItem).toEqual({ item: 0 });
		expect(result[0][1].pairedItem).toEqual({ item: 1 });
	});

	it('mixed-batch (item 0 invokeResponse, item 1 message) is rejected loudly', async () => {
		const node = new M365Agent();
		const sendResponse = vi.fn();
		const ctx = makeExecuteContext({
			inputItems: [{ a: 1 }, { b: 2 }],
			credentials: makeCredentials(),
			parameters: {
				'resource:0': 'invokeResponse',
				'operation:0': 'respond',
				'resource:1': 'message',
				'operation:1': 'send',
				responseShape: 'plain',
				statusCode: 200,
				body: '{}',
			},
		});
		(ctx as unknown as { sendResponse: typeof sendResponse }).sendResponse = sendResponse;
		await expect(node.execute.call(ctx as unknown as IExecuteFunctions)).rejects.toThrow(
			NodeOperationError,
		);
		await expect(node.execute.call(ctx as unknown as IExecuteFunctions)).rejects.toThrow(
			/Invoke Response must be the only resource/,
		);
		expect(sendResponse).not.toHaveBeenCalled();
	});
});
