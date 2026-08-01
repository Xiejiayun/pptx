import type {
  RasterImageContentType,
  SlideBackgroundImage,
} from '@pptx/model';
import {
  assertRasterImageContentType,
  resolveRasterImageSource,
  type RasterImageSource,
} from './raster-image-source.js';

const OPTION_KEYS = new Set(['contentType', 'signal']);

export interface SetSlideBackgroundImageOptions {
  readonly contentType?: RasterImageContentType;
  readonly signal?: AbortSignal;
}

interface NormalizedSetSlideBackgroundImageOptions {
  readonly contentType?: RasterImageContentType;
  readonly signal?: AbortSignal;
}

export async function resolveSlideBackgroundImage(
  source: RasterImageSource,
  options: SetSlideBackgroundImageOptions = {},
): Promise<SlideBackgroundImage> {
  const normalized = normalizeOptions(options);
  const resolved = await resolveRasterImageSource(source, normalized.signal);
  assertRasterImageContentType(normalized.contentType, resolved);
  return Object.freeze({
    kind: 'image',
    contentType: resolved.info.contentType,
    bytes: new Uint8Array(resolved.bytes),
  });
}

function normalizeOptions(options: unknown): NormalizedSetSlideBackgroundImageOptions {
  if (!options || typeof options !== 'object' || Array.isArray(options)) {
    throw new TypeError('Slide background image options must be an object');
  }
  const prototype = Object.getPrototypeOf(options);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError('Slide background image options must be an ordinary object');
  }
  const values = Object.create(null) as Record<string, unknown>;
  for (const key of Reflect.ownKeys(options)) {
    if (typeof key !== 'string' || !OPTION_KEYS.has(key)) {
      throw new TypeError(
        `Slide background image options contain unsupported property ${String(key)}`,
      );
    }
    const descriptor = Object.getOwnPropertyDescriptor(options, key);
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`Slide background image option ${key} must be a data property`);
    }
    values[key] = descriptor.value;
  }

  const contentType = normalizeContentType(values.contentType);
  const signal = normalizeSignal(values.signal);
  return Object.freeze({
    ...(contentType === undefined ? {} : { contentType }),
    ...(signal === undefined ? {} : { signal }),
  });
}

function normalizeContentType(value: unknown): RasterImageContentType | undefined {
  if (value === undefined) return undefined;
  if (value === 'image/png' || value === 'image/jpeg' || value === 'image/gif') {
    return value;
  }
  throw new TypeError('Slide background image contentType is unsupported');
}

function normalizeSignal(value: unknown): AbortSignal | undefined {
  if (value === undefined) return undefined;
  if (typeof AbortSignal !== 'undefined' && value instanceof AbortSignal) return value;
  throw new TypeError('Slide background image signal must be an AbortSignal');
}
