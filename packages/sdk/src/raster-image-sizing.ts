import type { ImageSourceRectangle } from '@pptx/model';
import type { RasterImageInfo } from './raster-image-source.js';

const INFO_KEYS = new Set(['contentType', 'width', 'height']);
const FRAME_KEYS = new Set(['type', 'width', 'height']);
const CROP_KEYS = new Set(['type', 'width', 'height', 'source']);
const SOURCE_KEYS = new Set(['x', 'y', 'width', 'height']);
const PERCENT_SCALE = 1_000;
const FULL_PERCENT = 100 * PERCENT_SCALE;
const INT32_MIN = -2_147_483_648;
const INT32_MAX = 2_147_483_647;

export interface RasterImageCropRegion {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

export type RasterImageSizing =
  | {
      readonly type: 'contain' | 'cover';
      readonly width: number;
      readonly height: number;
    }
  | {
      readonly type: 'crop';
      readonly width: number;
      readonly height: number;
      readonly source: RasterImageCropRegion;
    };

export interface RasterImageSizingResult {
  readonly width: number;
  readonly height: number;
  readonly sourceRectangle: Readonly<ImageSourceRectangle>;
}

export function normalizeRasterImageSizing(
  value: unknown,
): Readonly<RasterImageSizing> {
  const input = readClosedObject(value, CROP_KEYS, 'Raster image sizing');
  const type = input.type;
  if (type !== 'contain' && type !== 'cover' && type !== 'crop') {
    throw new TypeError('Raster image sizing type must be contain, cover, or crop');
  }
  const expectedKeys = type === 'crop' ? CROP_KEYS : FRAME_KEYS;
  requireExactKeys(input, expectedKeys, 'Raster image sizing');
  const width = positiveSafeInteger(input.width, 'Raster image sizing width');
  const height = positiveSafeInteger(input.height, 'Raster image sizing height');
  if (type !== 'crop') return Object.freeze({ type, width, height });

  const sourceInput = readClosedObject(
    input.source,
    SOURCE_KEYS,
    'Raster image sizing source',
  );
  requireExactKeys(sourceInput, SOURCE_KEYS, 'Raster image sizing source');
  const source = Object.freeze({
    x: nonNegativeFinite(sourceInput.x, 'Raster image sizing source x'),
    y: nonNegativeFinite(sourceInput.y, 'Raster image sizing source y'),
    width: positiveFinite(sourceInput.width, 'Raster image sizing source width'),
    height: positiveFinite(sourceInput.height, 'Raster image sizing source height'),
  });
  return Object.freeze({ type, width, height, source });
}

export function calculateRasterImageSizing(
  info: RasterImageInfo,
  sizing: RasterImageSizing,
): Readonly<RasterImageSizingResult> {
  const image = normalizeRasterImageInfo(info);
  const normalized = normalizeRasterImageSizing(sizing);
  const sourceRectangle = normalized.type === 'crop'
    ? calculateCrop(image, normalized.source)
    : calculateFit(image, normalized);
  return Object.freeze({
    width: normalized.width,
    height: normalized.height,
    sourceRectangle,
  });
}

function normalizeRasterImageInfo(value: unknown): RasterImageInfo {
  const input = readClosedObject(value, INFO_KEYS, 'Raster image info');
  requireExactKeys(input, INFO_KEYS, 'Raster image info');
  const contentType = input.contentType;
  if (contentType !== 'image/png' && contentType !== 'image/jpeg' && contentType !== 'image/gif') {
    throw new TypeError('Raster image info contentType is unsupported');
  }
  return Object.freeze({
    contentType,
    width: positiveSafeInteger(input.width, 'Raster image info width'),
    height: positiveSafeInteger(input.height, 'Raster image info height'),
  });
}

function calculateFit(
  image: RasterImageInfo,
  sizing: Extract<RasterImageSizing, { type: 'contain' | 'cover' }>,
): Readonly<ImageSourceRectangle> {
  const imageRatio = image.height / image.width;
  const frameRatio = sizing.height / sizing.width;
  if (frameRatio === imageRatio) return normalizeRectangle(0, 0, 0, 0);

  if (sizing.type === 'cover') {
    if (frameRatio > imageRatio) {
      const horizontal = 50 * (1 - sizing.width / (sizing.height / imageRatio));
      return normalizeRectangle(horizontal, 0, horizontal, 0);
    }
    const vertical = 50 * (1 - sizing.height / (sizing.width * imageRatio));
    return normalizeRectangle(0, vertical, 0, vertical);
  }

  if (frameRatio > imageRatio) {
    const vertical = 50 * (1 - sizing.height / (sizing.width * imageRatio));
    return normalizeRectangle(0, vertical, 0, vertical);
  }
  const horizontal = 50 * (1 - sizing.width / (sizing.height / imageRatio));
  return normalizeRectangle(horizontal, 0, horizontal, 0);
}

function calculateCrop(
  image: RasterImageInfo,
  source: Readonly<RasterImageCropRegion>,
): Readonly<ImageSourceRectangle> {
  if (source.x > image.width || source.width > image.width - source.x) {
    throw new RangeError('Raster image crop must fit within the intrinsic width');
  }
  if (source.y > image.height || source.height > image.height - source.y) {
    throw new RangeError('Raster image crop must fit within the intrinsic height');
  }
  return normalizeRectangle(
    100 * source.x / image.width,
    100 * source.y / image.height,
    100 * (image.width - source.x - source.width) / image.width,
    100 * (image.height - source.y - source.height) / image.height,
  );
}

function normalizeRectangle(
  left: number,
  top: number,
  right: number,
  bottom: number,
): Readonly<ImageSourceRectangle> {
  const raw = {
    left: normalizePercentage(left, 'left'),
    top: normalizePercentage(top, 'top'),
    right: normalizePercentage(right, 'right'),
    bottom: normalizePercentage(bottom, 'bottom'),
  };
  if (raw.left + raw.right >= FULL_PERCENT) {
    throw new RangeError('Raster image sizing must leave a positive source width');
  }
  if (raw.top + raw.bottom >= FULL_PERCENT) {
    throw new RangeError('Raster image sizing must leave a positive source height');
  }
  return Object.freeze({
    left: percentage(raw.left),
    top: percentage(raw.top),
    right: percentage(raw.right),
    bottom: percentage(raw.bottom),
  });
}

function normalizePercentage(value: number, edge: string): number {
  const raw = Math.round(value * PERCENT_SCALE);
  if (!Number.isSafeInteger(raw) || raw < INT32_MIN || raw > INT32_MAX) {
    throw new RangeError(`Raster image sizing ${edge} must fit the OOXML Int32 percentage range`);
  }
  if (raw >= FULL_PERCENT) {
    throw new RangeError(`Raster image sizing ${edge} must be less than 100 percent`);
  }
  return raw === 0 ? 0 : raw;
}

function percentage(raw: number): number {
  const value = raw / PERCENT_SCALE;
  return value === 0 ? 0 : value;
}

function readClosedObject(
  value: unknown,
  supportedKeys: ReadonlySet<string>,
  context: string,
): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${context} must be an object`);
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError(`${context} must be an ordinary object`);
  }
  const result = Object.create(null) as Record<string, unknown>;
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !supportedKeys.has(key)) {
      throw new TypeError(`${context} contains unsupported property ${String(key)}`);
    }
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`${context} ${key} must be a data property`);
    }
    result[key] = descriptor.value;
  }
  return result;
}

function requireExactKeys(
  value: Record<string, unknown>,
  keys: ReadonlySet<string>,
  context: string,
): void {
  for (const key of keys) {
    if (!Object.hasOwn(value, key)) throw new TypeError(`${context} ${key} is required`);
  }
  for (const key of Object.keys(value)) {
    if (!keys.has(key)) throw new TypeError(`${context} contains unsupported property ${key}`);
  }
}

function positiveSafeInteger(value: unknown, context: string): number {
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || value <= 0) {
    throw new RangeError(`${context} must be a positive safe integer`);
  }
  return value;
}

function nonNegativeFinite(value: unknown, context: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value) || value < 0) {
    throw new RangeError(`${context} must be a non-negative finite number`);
  }
  return value === 0 ? 0 : value;
}

function positiveFinite(value: unknown, context: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value) || value <= 0) {
    throw new RangeError(`${context} must be a positive finite number`);
  }
  return value;
}
