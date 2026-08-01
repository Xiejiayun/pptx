import { describe, expect, it } from 'vitest';
import { inches } from '@pptx/model';
import {
  calculateImageSizing,
  calculateRasterImageSizing,
  type ImageCropRegion,
  type ImageSizing,
  type ImageSizingResult,
  type RasterImageCropRegion,
  type RasterImageSizing,
  type RasterImageSizingResult,
} from './index.js';
import { normalizeRasterImageSizing } from './raster-image-sizing.js';

const LANDSCAPE = Object.freeze({
  contentType: 'image/png' as const,
  width: 1600,
  height: 900,
});
const PORTRAIT = Object.freeze({
  contentType: 'image/jpeg' as const,
  width: 900,
  height: 1600,
});
const SQUARE = Object.freeze({
  contentType: 'image/gif' as const,
  width: 1000,
  height: 1000,
});

describe('raster image sizing calculation', () => {
  it('publishes format-neutral aliases and sizes fractional SVG dimensions', () => {
    const source: ImageCropRegion = { x: 10.25, y: 5.5, width: 80.5, height: 40.25 };
    const sizing: ImageSizing = {
      type: 'crop',
      width: inches(4),
      height: inches(3),
      source,
    };
    const result: Readonly<ImageSizingResult> = calculateImageSizing(
      {
        contentType: 'image/svg+xml',
        width: 101.25,
        height: 50.625,
      },
      sizing,
    );

    expect(result).toEqual({
      width: inches(4),
      height: inches(3),
      sourceRectangle: {
        left: 10.123,
        top: 10.864,
        right: 10.37,
        bottom: 9.63,
      },
    });
    expect(Object.isFrozen(result)).toBe(true);
    expect(Object.isFrozen(result.sourceRectangle)).toBe(true);
  });

  it('matches raster contain, cover, and equal-ratio sizing for equivalent SVG ratios', () => {
    const cases: readonly ImageSizing[] = [
      { type: 'contain', width: inches(4), height: inches(3) },
      { type: 'cover', width: inches(4), height: inches(3) },
      { type: 'cover', width: inches(16), height: inches(9) },
    ];
    for (const sizing of cases) {
      const raster = calculateImageSizing(LANDSCAPE, sizing);
      const svg = calculateImageSizing({
        contentType: 'image/svg+xml',
        width: 17.6,
        height: 9.9,
      }, sizing);
      expect(svg).toEqual(raster);
    }
  });

  it('rejects unsafe generic image info while retaining strict raster validation', () => {
    const sizing = { type: 'contain' as const, width: 1, height: 1 };
    const invalid: unknown[] = [
      { contentType: 'image/svg+xml', width: 0, height: 1 },
      { contentType: 'image/svg+xml', width: Number.POSITIVE_INFINITY, height: 1 },
      { contentType: 'image/svg+xml', width: 1, height: Number.NaN },
      { contentType: 'image/svg+xml', width: 1, height: 1, extra: true },
      { contentType: 'image/svg', width: 1, height: 1 },
      Object.create({ contentType: 'image/svg+xml', width: 1, height: 1 }),
    ];
    for (const info of invalid) {
      expect(() => calculateImageSizing(info as never, sizing)).toThrow();
    }
    expect(() => calculateRasterImageSizing(
      { contentType: 'image/png', width: 1.5, height: 1 },
      sizing,
    )).toThrow(/safe integer/i);
    expect(calculateImageSizing(
      { contentType: 'image/svg+xml', width: 1.5, height: 1 },
      sizing,
    ).sourceRectangle).toEqual({ left: 0, top: -25, right: 0, bottom: -25 });
  });

  it('publishes the calculator and sizing types from the SDK root', () => {
    const source: RasterImageCropRegion = { x: 0, y: 0, width: 1600, height: 900 };
    const sizing: RasterImageSizing = {
      type: 'crop',
      width: inches(4),
      height: inches(3),
      source,
    };
    const result: Readonly<RasterImageSizingResult> = calculateRasterImageSizing(
      LANDSCAPE,
      sizing,
    );

    expect(result).toEqual({
      width: inches(4),
      height: inches(3),
      sourceRectangle: { left: 0, top: 0, right: 0, bottom: 0 },
    });
  });

  it('calculates exact cover branches and equal-ratio zero state', () => {
    expect(calculateRasterImageSizing(LANDSCAPE, {
      type: 'cover',
      width: inches(4),
      height: inches(3),
    })).toEqual({
      width: inches(4),
      height: inches(3),
      sourceRectangle: { left: 12.5, top: 0, right: 12.5, bottom: 0 },
    });
    expect(calculateRasterImageSizing(PORTRAIT, {
      type: 'cover',
      width: inches(4),
      height: inches(3),
    })).toEqual({
      width: inches(4),
      height: inches(3),
      sourceRectangle: { left: 0, top: 28.906, right: 0, bottom: 28.906 },
    });
    expect(calculateRasterImageSizing(LANDSCAPE, {
      type: 'cover',
      width: inches(16),
      height: inches(9),
    }).sourceRectangle).toEqual({ left: 0, top: 0, right: 0, bottom: 0 });
  });

  it('calculates exact contain branches as symmetric negative extensions', () => {
    expect(calculateRasterImageSizing(LANDSCAPE, {
      type: 'contain',
      width: inches(4),
      height: inches(3),
    })).toEqual({
      width: inches(4),
      height: inches(3),
      sourceRectangle: { left: 0, top: -16.667, right: 0, bottom: -16.667 },
    });
    expect(calculateRasterImageSizing(PORTRAIT, {
      type: 'contain',
      width: inches(4),
      height: inches(3),
    })).toEqual({
      width: inches(4),
      height: inches(3),
      sourceRectangle: { left: -68.519, top: 0, right: -68.519, bottom: 0 },
    });
    expect(calculateRasterImageSizing(SQUARE, {
      type: 'contain',
      width: inches(2),
      height: inches(2),
    }).sourceRectangle).toEqual({ left: 0, top: 0, right: 0, bottom: 0 });
  });

  it('maps full, centered, fractional, and edge-aligned pixel crop regions', () => {
    expect(calculateRasterImageSizing(LANDSCAPE, {
      type: 'crop',
      width: inches(3),
      height: inches(2),
      source: { x: 0, y: 0, width: 1600, height: 900 },
    }).sourceRectangle).toEqual({ left: 0, top: 0, right: 0, bottom: 0 });
    expect(calculateRasterImageSizing(LANDSCAPE, {
      type: 'crop',
      width: inches(3),
      height: inches(2),
      source: { x: 400, y: 225, width: 800, height: 450 },
    })).toEqual({
      width: inches(3),
      height: inches(2),
      sourceRectangle: { left: 25, top: 25, right: 25, bottom: 25 },
    });
    expect(calculateRasterImageSizing(
      { contentType: 'image/png', width: 100, height: 80 },
      {
        type: 'crop',
        width: inches(5),
        height: inches(1),
        source: { x: 0.5, y: 1.25, width: 50, height: 40.5 },
      },
    ).sourceRectangle).toEqual({
      left: 0.5,
      top: 1.563,
      right: 49.5,
      bottom: 47.813,
    });
    expect(calculateRasterImageSizing(
      { contentType: 'image/png', width: 100, height: 80 },
      {
        type: 'crop',
        width: inches(1),
        height: inches(1),
        source: { x: 75, y: 60, width: 25, height: 20 },
      },
    ).sourceRectangle).toEqual({ left: 75, top: 75, right: 0, bottom: 0 });
  });

  it('normalizes descriptor-safe detached sizing values before asynchronous use', () => {
    const source = { x: 10.25, y: 20.5, width: 100, height: 200 };
    const input: RasterImageSizing = {
      type: 'crop',
      width: inches(3),
      height: inches(2),
      source,
    };
    const normalized = normalizeRasterImageSizing(input);
    source.x = 999;
    (input as { width: number }).width = 1;

    expect(normalized).toEqual({
      type: 'crop',
      width: inches(3),
      height: inches(2),
      source: { x: 10.25, y: 20.5, width: 100, height: 200 },
    });
    expect(Object.isFrozen(normalized)).toBe(true);
    expect(Object.isFrozen((normalized as Extract<RasterImageSizing, { type: 'crop' }>).source))
      .toBe(true);

    const result = calculateRasterImageSizing(
      { contentType: 'image/png', width: 200, height: 300 },
      normalized,
    );
    expect(Object.isFrozen(result)).toBe(true);
    expect(Object.isFrozen(result.sourceRectangle)).toBe(true);
  });

  it('accepts ordinary and null-prototype data containers', () => {
    const source = Object.assign(Object.create(null) as RasterImageCropRegion, {
      x: 400,
      y: 225,
      width: 800,
      height: 450,
    });
    const sizing = Object.assign(Object.create(null) as RasterImageSizing, {
      type: 'crop' as const,
      width: inches(3),
      height: inches(2),
      source,
    });
    const info = Object.assign(Object.create(null) as typeof LANDSCAPE, LANDSCAPE);

    expect(calculateRasterImageSizing(info, sizing).sourceRectangle).toEqual({
      left: 25,
      top: 25,
      right: 25,
      bottom: 25,
    });
  });

  it('rejects unsafe sizing containers and fields without invoking accessors', () => {
    class Sizing {
      type = 'cover';
      width = 1;
      height = 1;
    }
    let reads = 0;
    const accessor = Object.defineProperty({ type: 'cover', height: 1 }, 'width', {
      enumerable: true,
      get() {
        reads += 1;
        return 1;
      },
    });
    const inherited = Object.create({ type: 'cover', width: 1, height: 1 });
    const symbol = { type: 'cover', width: 1, height: 1, [Symbol('unsafe')]: true };
    const sourceAccessor = Object.defineProperty(
      { x: 0, y: 0, height: 1 },
      'width',
      {
        enumerable: true,
        get() {
          reads += 1;
          return 1;
        },
      },
    );
    class Source {
      x = 0;
      y = 0;
      width = 1;
      height = 1;
    }
    const validCrop = {
      type: 'crop',
      width: 1,
      height: 1,
      source: { x: 0, y: 0, width: 1, height: 1 },
    };
    const invalid = [
      null,
      undefined,
      [],
      new Sizing(),
      accessor,
      inherited,
      symbol,
      {},
      { type: 'fit', width: 1, height: 1 },
      { type: 'cover', width: 1, height: 1, unknown: true },
      { type: 'cover', width: 0, height: 1 },
      { type: 'cover', width: 1.5, height: 1 },
      { type: 'contain', width: 1, height: Number.POSITIVE_INFINITY },
      { type: 'crop', width: 1, height: 1 },
      { ...validCrop, source: null },
      { ...validCrop, source: new Source() },
      { ...validCrop, source: sourceAccessor },
      { ...validCrop, source: Object.create({ x: 0, y: 0, width: 1, height: 1 }) },
      { ...validCrop, source: { x: 0, y: 0, width: 1, height: 1, [Symbol('unsafe')]: true } },
      { ...validCrop, source: { x: 0, y: 0, width: 1 } },
      { ...validCrop, source: { x: 0, y: 0, width: 1, height: 1, unknown: true } },
      { ...validCrop, source: { x: -1, y: 0, width: 1, height: 1 } },
      { ...validCrop, source: { x: 0, y: Number.NaN, width: 1, height: 1 } },
      { ...validCrop, source: { x: 0, y: 0, width: 0, height: 1 } },
    ];

    for (const [index, value] of invalid.entries()) {
      expect(() => normalizeRasterImageSizing(value), `invalid sizing ${index}`).toThrow();
    }
    expect(reads).toBe(0);
  });

  it('rejects invalid image info, out-of-bounds crops, and percentage overflow', () => {
    const cover = { type: 'cover', width: 1, height: 1 } as const;
    let reads = 0;
    const accessor = Object.defineProperty(
      { contentType: 'image/png', height: 1 },
      'width',
      {
        enumerable: true,
        get() {
          reads += 1;
          return 1;
        },
      },
    );
    class Info {
      contentType = 'image/png';
      width = 1;
      height = 1;
    }
    for (const info of [
      null,
      {},
      new Info(),
      accessor,
      Object.create({ contentType: 'image/png', width: 1, height: 1 }),
      { contentType: 'image/png', width: 1, height: 1, [Symbol('unsafe')]: true },
      { contentType: 'image/svg+xml', width: 1, height: 1 },
      { contentType: 'image/png', width: 0, height: 1 },
      { contentType: 'image/png', width: 1.5, height: 1 },
      { contentType: 'image/png', width: 1, height: 1, unknown: true },
    ]) {
      expect(() => calculateRasterImageSizing(info as never, cover)).toThrow();
    }
    expect(reads).toBe(0);

    for (const source of [
      { x: 101, y: 0, width: 1, height: 1 },
      { x: 90, y: 0, width: 11, height: 1 },
      { x: 0, y: 81, width: 1, height: 1 },
      { x: 0, y: 70, width: 1, height: 11 },
    ]) {
      expect(() => calculateRasterImageSizing(
        { contentType: 'image/png', width: 100, height: 80 },
        { type: 'crop', width: 1, height: 1, source },
      )).toThrow(RangeError);
    }

    expect(() => calculateRasterImageSizing(
      { contentType: 'image/png', width: Number.MAX_SAFE_INTEGER, height: 1 },
      { type: 'contain', width: 1, height: Number.MAX_SAFE_INTEGER },
    )).toThrow(RangeError);
  });
});
