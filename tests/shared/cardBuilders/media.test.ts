import { describe, it, expect } from 'vitest';
import { mediaField, buildMedia } from '../../../shared/cardBuilders/media';

describe('mediaField', () => {
  it('is a fixedCollection with multipleValues', () => {
    const field = mediaField();
    expect(field.type).toBe('fixedCollection');
    expect(field.typeOptions).toEqual({ multipleValues: true });
    expect(field.name).toBe('media');
    expect(field.displayName).toBe('Media URLs');
    expect(field.default).toEqual({});
  });

  it('media row has url and options fields', () => {
    const field = mediaField();
    const row = (field.options as Array<{ values: Array<{ name: string }> }>)[0];
    expect(row.values.map((v) => v.name)).toEqual(['url', 'options']);
  });
});

describe('buildMedia', () => {
  it('returns [] when empty', () => {
    expect(buildMedia(undefined)).toEqual([]);
    expect(buildMedia({ mediaItem: [] })).toEqual([]);
  });

  it('maps url directly; profile from Options.profile', () => {
    const out = buildMedia({
      mediaItem: [
        { url: 'https://a' },
        { url: 'https://b', options: { profile: 'hd' } },
      ],
    });
    expect(out).toEqual([{ url: 'https://a' }, { url: 'https://b', profile: 'hd' }]);
  });
});
