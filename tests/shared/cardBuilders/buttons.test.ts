import { describe, it, expect } from 'vitest';
import { buttonsField, buildButtons } from '../../../shared/cardBuilders/buttons';

describe('buttonsField', () => {
  it('is a fixedCollection with multipleValues', () => {
    const field = buttonsField();
    expect(field.type).toBe('fixedCollection');
    expect(field.typeOptions).toEqual({ multipleValues: true });
    expect(field.name).toBe('buttons');
    expect(field.displayName).toBe('Buttons');
    expect(field.default).toEqual({});
  });

  it('has a single "button" row with the cardActionFields shape', () => {
    const field = buttonsField();
    const collection = field.options as Array<{
      name: string;
      values: Array<{ name: string }>;
    }>;
    expect(collection).toHaveLength(1);
    expect(collection[0].name).toBe('button');
    expect(collection[0].values.map((v) => v.name)).toEqual(['type', 'title', 'value', 'options']);
  });
});

describe('buildButtons', () => {
  it('returns [] when the collection is absent', () => {
    expect(buildButtons(undefined)).toEqual([]);
    expect(buildButtons({})).toEqual([]);
    expect(buildButtons({ button: [] })).toEqual([]);
  });

  it('maps each button row via buildActions', () => {
    const out = buildButtons({
      button: [
        { type: 'openUrl', title: 'Open', value: 'https://x' },
        { type: 'imBack', title: 'Yes', value: 'Yes' },
      ],
    });
    expect(out).toEqual([
      { type: 'openUrl', title: 'Open', value: 'https://x' },
      { type: 'imBack', title: 'Yes', value: 'Yes' },
    ]);
  });
});
