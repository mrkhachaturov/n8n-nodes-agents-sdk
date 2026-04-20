import { describe, it, expect } from 'vitest';
import { imagesField, buildImages } from '../../../shared/cardBuilders/images';

describe('imagesField', () => {
	it('is a fixedCollection with multipleValues, display name "Images"', () => {
		const field = imagesField();
		expect(field.type).toBe('fixedCollection');
		expect(field.typeOptions).toEqual({ multipleValues: true });
		expect(field.name).toBe('images');
		expect(field.displayName).toBe('Images');
		expect(field.default).toEqual({});
	});

	it('image row has url, alt, options fields', () => {
		const field = imagesField();
		const row = (field.options as Array<{ values: Array<{ name: string }> }>)[0];
		expect(row.values.map((v) => v.name)).toEqual(['url', 'alt', 'options']);
	});

	it('Tap Action inside Options is a singular fixedCollection (not collection)', () => {
		// Why: n8n's frontend validator (packages/workflow/src/node-helpers.ts:1540)
		// walks every direct child of a `collection` unconditionally — no
		// `value === undefined` guard. Wrapping cardActionFields() (which has
		// `required: true` on type/title/value) in a singular `fixedCollection`
		// makes the validator skip the inner walk until the user opens
		// "Add Tap Action".
		const field = imagesField();
		const row = (
			field.options as Array<{
				values: Array<{
					name: string;
					type?: string;
					typeOptions?: { multipleValues?: boolean };
					options?: Array<{
						name: string;
						type: string;
						typeOptions?: { multipleValues?: boolean };
						options?: Array<{ name: string; values?: Array<{ name: string }> }>;
					}>;
				}>;
			}>
		)[0];
		const options = row.values.find((v) => v.name === 'options')!;
		const tap = options.options!.find((o) => o.name === 'tapAction')!;
		expect(tap.type).toBe('fixedCollection');
		expect(tap.typeOptions).toEqual({ multipleValues: false });
		// One named inner group ("action") wraps cardActionFields().
		const actionGroup = tap.options![0];
		expect(actionGroup.name).toBe('action');
		expect(actionGroup.values!.map((v) => v.name)).toEqual([
			'type',
			'title',
			'value',
			'options',
		]);
	});
});

describe('buildImages', () => {
	it('returns [] when empty', () => {
		expect(buildImages(undefined)).toEqual([]);
		expect(buildImages({})).toEqual([]);
		expect(buildImages({ image: [] })).toEqual([]);
	});

	it('maps url and alt directly; omits alt when empty', () => {
		const out = buildImages({
			image: [
				{ url: 'https://a', alt: 'hello' },
				{ url: 'https://b', alt: '' },
			],
		});
		expect(out).toEqual([{ url: 'https://a', alt: 'hello' }, { url: 'https://b' }]);
	});

	it('extracts tap from Options.tapAction.action (the fixedCollection-wrapped shape)', () => {
		const out = buildImages({
			image: [
				{
					url: 'https://a',
					alt: 'alt',
					options: {
						tapAction: {
							action: { type: 'openUrl', title: 'Open', value: 'https://x' },
						},
					},
				},
			],
		});
		expect(out[0].tap).toEqual({ type: 'openUrl', title: 'Open', value: 'https://x' });
	});

	it('omits tap when Options.tapAction is absent, empty, or missing the action wrapper', () => {
		const empty = buildImages({ image: [{ url: 'u', alt: 'a', options: {} }] });
		expect(empty[0]).not.toHaveProperty('tap');
		const missing = buildImages({ image: [{ url: 'u', alt: 'a' }] });
		expect(missing[0]).not.toHaveProperty('tap');
		const wrapperPresentNoAction = buildImages({
			image: [{ url: 'u', alt: 'a', options: { tapAction: {} } }],
		});
		expect(wrapperPresentNoAction[0]).not.toHaveProperty('tap');
	});
});
