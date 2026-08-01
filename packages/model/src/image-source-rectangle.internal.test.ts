import { describe, expect, it } from 'vitest';
import {
  normalizeImageSourceRectangle,
  renderImageSourceRectangle,
} from './image-source-rectangle.internal.js';

describe('image source rectangle normalization', () => {
  it('normalizes signed percentages to detached frozen thousandths', () => {
    const input = {
      left: 12.3456,
      top: -20.0004,
      right: 1,
      bottom: -0,
    };
    const normalized = normalizeImageSourceRectangle(input, 'Image source rectangle');
    input.left = 50;
    input.top = 50;

    expect(normalized).toEqual({
      left: 12.346,
      top: -20,
      right: 1,
      bottom: 0,
    });
    expect(Object.isFrozen(normalized)).toBe(true);
  });

  it('accepts null-prototype explicit zero and negative contain values', () => {
    const value = Object.create(null) as Record<string, unknown>;
    value.left = -25;
    value.top = 0;
    value.right = -25;
    value.bottom = 0;

    expect(normalizeImageSourceRectangle(value, 'Image source rectangle')).toEqual({
      left: -25,
      top: 0,
      right: -25,
      bottom: 0,
    });
    expect(normalizeImageSourceRectangle({
      left: 0,
      top: 0,
      right: 0,
      bottom: 0,
    }, 'Image source rectangle')).toEqual({
      left: 0,
      top: 0,
      right: 0,
      bottom: 0,
    });
  });

  it('rejects unsafe containers and descriptors without invoking accessors', () => {
    class Rectangle {
      left = 0;
      top = 0;
      right = 0;
      bottom = 0;
    }
    let reads = 0;
    const accessor = Object.defineProperty({ top: 0, right: 0, bottom: 0 }, 'left', {
      enumerable: true,
      get() {
        reads += 1;
        return 0;
      },
    });
    const inherited = Object.create({ left: 0, top: 0, right: 0, bottom: 0 });
    const symbol = {
      left: 0,
      top: 0,
      right: 0,
      bottom: 0,
      [Symbol('unsafe')]: true,
    };

    for (const value of [null, undefined, [], new Rectangle(), accessor, inherited, symbol]) {
      expect(() => normalizeImageSourceRectangle(value, 'Image source rectangle')).toThrow(
        TypeError,
      );
    }
    expect(reads).toBe(0);
  });

  it('rejects missing, unknown, and invalid edge values', () => {
    const valid = { left: 0, top: 0, right: 0, bottom: 0 };
    const cases = [
      {},
      { top: 0, right: 0, bottom: 0 },
      { ...valid, unknown: true },
      { ...valid, left: '0' },
      { ...valid, left: Number.NaN },
      { ...valid, left: Number.NEGATIVE_INFINITY },
      { ...valid, left: Number.POSITIVE_INFINITY },
      { ...valid, left: -2_147_483.649 },
      { ...valid, left: 100 },
      { ...valid, left: 60, right: 40 },
      { ...valid, top: 75, bottom: 25 },
    ];

    for (const value of cases) {
      expect(() => normalizeImageSourceRectangle(value, 'Image source rectangle')).toThrow();
    }
  });

  it('renders canonical DrawingML integer percentages', () => {
    const value = normalizeImageSourceRectangle({
      left: 12.3456,
      top: -20.0004,
      right: 1,
      bottom: 0,
    }, 'Image source rectangle');

    expect(renderImageSourceRectangle(value)).toBe(
      '<a:srcRect l="12346" t="-20000" r="1000" b="0"/>',
    );
    expect(renderImageSourceRectangle(value, 'd')).toBe(
      '<d:srcRect l="12346" t="-20000" r="1000" b="0"/>',
    );
    expect(renderImageSourceRectangle(value, '')).toBe(
      '<srcRect l="12346" t="-20000" r="1000" b="0"/>',
    );
  });
});
