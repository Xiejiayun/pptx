import { describe, expect, it } from 'vitest';
import { SCHEME_COLORS as INDEX_SCHEME_COLORS } from './index.js';
import { SCHEME_COLORS, type SchemeColor } from './scheme-color.js';

const EXPECTED = {
  text1: 'tx1',
  text2: 'tx2',
  background1: 'bg1',
  background2: 'bg2',
  accent1: 'accent1',
  accent2: 'accent2',
  accent3: 'accent3',
  accent4: 'accent4',
  accent5: 'accent5',
  accent6: 'accent6',
} as const;

describe('SCHEME_COLORS', () => {
  it('publishes the complete frozen PptxGenJS helper mapping', () => {
    expect(SCHEME_COLORS).toBe(INDEX_SCHEME_COLORS);
    expect(Object.entries(SCHEME_COLORS)).toEqual(Object.entries(EXPECTED));
    expect(new Set(Object.values(SCHEME_COLORS))).toHaveLength(10);
    expect(Object.isFrozen(SCHEME_COLORS)).toBe(true);
    expect(() => {
      (SCHEME_COLORS as unknown as { accent1: string }).accent1 = 'changed';
    }).toThrow(TypeError);
    expect(() => Object.defineProperty(SCHEME_COLORS, 'accent7', {
      value: 'accent7',
    })).toThrow(TypeError);
    expect(() => {
      delete (SCHEME_COLORS as unknown as { accent6?: string }).accent6;
    }).toThrow(TypeError);
    expect(SCHEME_COLORS).toEqual(EXPECTED);
  });

  it('derives SchemeColor from the helper values', () => {
    const values: readonly SchemeColor[] = Object.values(SCHEME_COLORS);
    expect(values).toEqual(Object.values(EXPECTED));
    if (false) {
      // @ts-expect-error helper key is not a helper value
      const key: SchemeColor = 'text1';
      // @ts-expect-error extended DrawingML token is outside this PptxGenJS helper
      const extended: SchemeColor = 'hlink';
      // @ts-expect-error raw sRGB is not a scheme helper value
      const srgb: SchemeColor = 'FF0000';
      // @ts-expect-error helper keys are closed
      SCHEME_COLORS.text3;
      void [key, extended, srgb];
    }
  });
});
