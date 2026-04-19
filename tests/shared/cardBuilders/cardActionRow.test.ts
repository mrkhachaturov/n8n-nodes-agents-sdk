import { describe, it, expect } from 'vitest';
import type { INodeProperties } from 'n8n-workflow';
import {
	cardActionFields,
	buildActions,
	CardBuildError,
} from '../../../shared/cardBuilders/cardActionRow';

describe('cardActionFields', () => {
	const fields: INodeProperties[] = cardActionFields();

	it('returns exactly four top-level fields (Type, Title, Value, Options)', () => {
		expect(fields.map((f) => f.name)).toEqual(['type', 'title', 'value', 'options']);
	});

	it('Type is a required options field with 11 alphabetized ActionTypes', () => {
		const typeField = fields.find((f) => f.name === 'type')!;
		expect(typeField.type).toBe('options');
		expect(typeField.required).toBe(true);
		expect(typeField.noDataExpression).toBe(true);
		expect(typeField.default).toBe('openUrl');
		const names = (typeField.options as { name: string; value: string }[]).map((o) => o.name);
		expect(names).toEqual([
			'Call',
			'Download File',
			'IM Back',
			'Message Back',
			'Open App',
			'Open URL',
			'Play Audio',
			'Play Video',
			'Post Back',
			'Show Image',
			'Sign In',
		]);
		const values = (typeField.options as { value: string }[]).map((o) => o.value);
		expect(values).toEqual([
			'call',
			'downloadFile',
			'imBack',
			'messageBack',
			'openApp',
			'openUrl',
			'playAudio',
			'playVideo',
			'postBack',
			'showImage',
			'signin',
		]);
	});

	it('Title and Value are required strings with empty defaults', () => {
		const title = fields.find((f) => f.name === 'title')!;
		const value = fields.find((f) => f.name === 'value')!;
		expect(title.type).toBe('string');
		expect(title.required).toBe(true);
		expect(title.default).toBe('');
		expect(value.type).toBe('string');
		expect(value.required).toBe(true);
		expect(value.default).toBe('');
	});

	it('Options collection has alphabetized entries including Value JSON and Channel Data', () => {
		const options = fields.find((f) => f.name === 'options')!;
		expect(options.type).toBe('collection');
		expect(options.default).toEqual({});
		const opts = options.options as INodeProperties[];
		expect(opts.map((o) => o.name)).toEqual([
			'channelData',
			'displayText',
			'image',
			'imageAltText',
			'text',
			'valueJson',
		]);
		const valueJson = opts.find((o) => o.name === 'valueJson')!;
		expect(valueJson.type).toBe('json');
		expect(valueJson.default).toBe('={}');
		const channelData = opts.find((o) => o.name === 'channelData')!;
		expect(channelData.type).toBe('json');
		expect(channelData.default).toBe('={}');
	});
});

describe('buildActions', () => {
	it('maps the basic row shape to CardAction', () => {
		const out = buildActions([{ type: 'openUrl', title: 'Open', value: 'https://x' }]);
		expect(out).toEqual([{ type: 'openUrl', title: 'Open', value: 'https://x' }]);
	});

	it('returns [] for empty input', () => {
		expect(buildActions([])).toEqual([]);
		expect(buildActions(undefined)).toEqual([]);
	});

	it('Options.valueJson (non-empty) overrides Value and is parsed as JSON', () => {
		const out = buildActions([
			{
				type: 'messageBack',
				title: 'Reply',
				value: '',
				options: { valueJson: '{"kind":"accept","id":42}' },
			},
		]);
		expect(out[0].value).toEqual({ kind: 'accept', id: 42 });
	});

	it('Options.valueJson accepts a JSON array payload', () => {
		const out = buildActions([
			{
				type: 'messageBack',
				title: 'Bulk',
				value: '',
				options: { valueJson: '[1,2,3]' },
			},
		]);
		expect(out[0].value).toEqual([1, 2, 3]);
	});

	it('Options.valueJson accepts a JSON scalar payload', () => {
		const outNumber = buildActions([
			{ type: 'postBack', title: 'Count', value: '', options: { valueJson: '42' } },
		]);
		expect(outNumber[0].value).toBe(42);

		const outBoolean = buildActions([
			{ type: 'postBack', title: 'Flag', value: '', options: { valueJson: 'true' } },
		]);
		expect(outBoolean[0].value).toBe(true);

		const outString = buildActions([
			{ type: 'postBack', title: 'Tag', value: '', options: { valueJson: '"hello"' } },
		]);
		expect(outString[0].value).toBe('hello');
	});

	it('Options.channelData accepts array or scalar payloads alongside objects', () => {
		const outArr = buildActions([
			{
				type: 'openUrl',
				title: 'X',
				value: 'https://x',
				options: { channelData: '["a","b"]' },
			},
		]);
		expect(outArr[0].channelData).toEqual(['a', 'b']);
		const outNum = buildActions([
			{
				type: 'openUrl',
				title: 'X',
				value: 'https://x',
				options: { channelData: '7' },
			},
		]);
		expect(outNum[0].channelData).toBe(7);
	});

	it('Options.valueJson passes through a pre-parsed non-string payload untouched', () => {
		const out = buildActions([
			{
				type: 'messageBack',
				title: 'X',
				value: '',
				options: { valueJson: { already: 'parsed' } as unknown },
			},
		]);
		expect(out[0].value).toEqual({ already: 'parsed' });
	});

	it('Options.valueJson when empty-string or "={}" keeps the Value string', () => {
		const emptyString = buildActions([
			{ type: 'postBack', title: 'B', value: 'raw', options: { valueJson: '' } },
		]);
		expect(emptyString[0].value).toBe('raw');
		const defaultExpr = buildActions([
			{ type: 'postBack', title: 'B', value: 'raw', options: { valueJson: '={}' } },
		]);
		expect(defaultExpr[0].value).toBe('raw');
	});

	it('maps Options.channelData through JSON parse', () => {
		const out = buildActions([
			{
				type: 'openUrl',
				title: 'X',
				value: 'https://x',
				options: { channelData: '{"teamsChannelId":"c1"}' },
			},
		]);
		expect(out[0].channelData).toEqual({ teamsChannelId: 'c1' });
	});

	it('maps image, imageAltText, text, displayText 1:1', () => {
		const out = buildActions([
			{
				type: 'imBack',
				title: 'Yes',
				value: 'Yes',
				options: {
					image: 'https://img',
					imageAltText: 'alt',
					text: 't',
					displayText: 'd',
				},
			},
		]);
		expect(out[0]).toMatchObject({
			type: 'imBack',
			title: 'Yes',
			value: 'Yes',
			image: 'https://img',
			imageAltText: 'alt',
			text: 't',
			displayText: 'd',
		});
	});

	it('throws CardBuildError on invalid valueJson (dispatchers classify as user config)', () => {
		expect(() =>
			buildActions([
				{ type: 'messageBack', title: 'X', value: '', options: { valueJson: '{not-json' } },
			]),
		).toThrow(CardBuildError);
		expect(() =>
			buildActions([
				{ type: 'messageBack', title: 'X', value: '', options: { valueJson: '{not-json' } },
			]),
		).toThrow(/Invalid JSON in Value JSON/);
	});

	it('throws CardBuildError on invalid channelData as well', () => {
		expect(() =>
			buildActions([
				{ type: 'openUrl', title: 'X', value: 'https://x', options: { channelData: '{bad' } },
			]),
		).toThrow(CardBuildError);
	});
});
