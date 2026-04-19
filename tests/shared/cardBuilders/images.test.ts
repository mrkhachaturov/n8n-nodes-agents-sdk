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

  it('Tap Action inside Options is a collection (singleton), not fixedCollection', () => {
    const field = imagesField();
    const row = (field.options as Array<{ values: Array<{ name: string; type?: string; options?: Array<{ name: string; type: string }> }> }>)[0];
    const options = row.values.find((v) => v.name === 'options')!;
    const tap = options.options!.find((o) => o.name === 'tapAction')!;
    expect(tap.type).toBe('collection');
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
      image: [{ url: 'https://a', alt: 'hello' }, { url: 'https://b', alt: '' }],
    });
    expect(out).toEqual([
      { url: 'https://a', alt: 'hello' },
      { url: 'https://b' },
    ]);
  });

  it('extracts tap singleton from Options.tapAction', () => {
    const out = buildImages({
      image: [
        {
          url: 'https://a',
          alt: 'alt',
          options: {
            tapAction: { type: 'openUrl', title: 'Open', value: 'https://x' },
          },
        },
      ],
    });
    expect(out[0].tap).toEqual({ type: 'openUrl', title: 'Open', value: 'https://x' });
  });

  it('omits tap when Options.tapAction is absent or empty', () => {
    const empty = buildImages({ image: [{ url: 'u', alt: 'a', options: {} }] });
    expect(empty[0]).not.toHaveProperty('tap');
    const missing = buildImages({ image: [{ url: 'u', alt: 'a' }] });
    expect(missing[0]).not.toHaveProperty('tap');
  });
});
