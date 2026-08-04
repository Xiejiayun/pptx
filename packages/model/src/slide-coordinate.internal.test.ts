import { describe, expect, it } from 'vitest';
import { resolveSlideCoordinate } from './slide-coordinate.internal.js';
import { inches, type Emu, type SlideSize } from './units.js';

const ZERO = 0 as Emu;
const SIZE: Readonly<SlideSize> = Object.freeze({
  width: inches(10),
  height: inches(8),
});

describe('slide-relative coordinates', () => {
  it('resolves strict percentages on the selected slide axis', () => {
    expect(resolveSlideCoordinate('10%', 'horizontal', SIZE, inches(1), 'Shape x'))
      .toBe(inches(1));
    expect(resolveSlideCoordinate('20%', 'vertical', SIZE, inches(1), 'Shape y'))
      .toBe(inches(1.6));
    expect(resolveSlideCoordinate('1.25e1%', 'horizontal', SIZE, ZERO, 'Shape width'))
      .toBe(inches(1.25));
    expect(resolveSlideCoordinate('-25%', 'horizontal', SIZE, ZERO, 'Shape x'))
      .toBe(inches(-2.5));
    expect(resolveSlideCoordinate('125%', 'vertical', SIZE, ZERO, 'Shape y'))
      .toBe(inches(10));
    expect(Object.is(resolveSlideCoordinate('-0%', 'vertical', SIZE, ZERO, 'Shape y'), -0))
      .toBe(false);
  });

  it('preserves absolute EMU inputs and caller defaults', () => {
    expect(resolveSlideCoordinate(inches(2), 'horizontal', SIZE, ZERO, 'Shape x'))
      .toBe(inches(2));
    expect(resolveSlideCoordinate(1.6, 'horizontal', SIZE, ZERO, 'Shape x')).toBe(2);
    expect(resolveSlideCoordinate(undefined, 'horizontal', SIZE, inches(3), 'Shape x'))
      .toBe(inches(3));
  });

  it('rejects malformed, non-finite, context-free, and unsafe inputs', () => {
    for (const value of [
      '',
      '10',
      '%',
      ' 10%',
      '10% ',
      '10%%',
      '10%garbage',
      'NaN%',
      'Infinity%',
      {},
      [],
      true,
      null,
      Number.NaN,
      Number.POSITIVE_INFINITY,
    ]) {
      expect(() => resolveSlideCoordinate(
        value,
        'horizontal',
        SIZE,
        ZERO,
        'Shape x',
      ), String(value)).toThrow(TypeError);
    }

    expect(() => resolveSlideCoordinate(
      '10%',
      'horizontal',
      undefined,
      ZERO,
      'Shape x',
    )).toThrow(/slide size/u);
    expect(() => resolveSlideCoordinate(
      '1e20%',
      'horizontal',
      SIZE,
      ZERO,
      'Shape x',
    )).toThrow(RangeError);
    expect(() => resolveSlideCoordinate(
      Number.MAX_SAFE_INTEGER + 1,
      'horizontal',
      SIZE,
      ZERO,
      'Shape x',
    )).toThrow(RangeError);
  });
});
