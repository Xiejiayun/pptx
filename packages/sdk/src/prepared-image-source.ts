import type {
  AddSvgImageOptions,
  ImageModel,
  PlaceholderSelector,
  SlideModel,
} from '@pptx/model';
import {
  assertImageContentType,
  normalizeAddImageSourceOptions,
  resolveImageSource,
  type AddImageSourceOptions,
  type ImageSource,
  type NormalizedAddImageSourceOptions,
  type ResolvedImageSource,
} from './raster-image-source.js';
import {
  calculateImageSizing,
  type ImageSizing,
} from './raster-image-sizing.js';
import { resolveSvgFallback } from './svg-image-fallback.js';

export interface PreparedImageSource {
  readonly resolved: Readonly<ResolvedImageSource>;
  readonly options: Readonly<NormalizedAddImageSourceOptions>;
  readonly fallbackPngBytes?: Uint8Array;
}

export async function prepareImageSource(
  source: ImageSource,
  options: AddImageSourceOptions = {},
): Promise<Readonly<PreparedImageSource>> {
  const normalized = normalizeAddImageSourceOptions(options);
  const resolved = await resolveImageSource(source, normalized.signal);
  assertImageContentType(normalized.contentType, resolved);
  if (resolved.info.contentType !== 'image/svg+xml') {
    if (normalized.fallback !== undefined) {
      throw new TypeError('fallback is only valid for SVG images');
    }
    return Object.freeze({ resolved, options: normalized });
  }
  return Object.freeze({
    resolved,
    options: normalized,
    fallbackPngBytes: await resolveSvgFallback(
      resolved,
      normalized.fallback,
      normalized.signal,
    ),
  });
}

export function commitPreparedImage(
  slide: SlideModel,
  prepared: Readonly<PreparedImageSource>,
): ImageModel {
  const { resolved, options } = prepared;
  const placement = options.sizing === undefined
    ? undefined
    : calculateImageSizing(
        resolved.info,
        imageSizingForPlaceholder(
          slide,
          options.imageOptions.placeholder,
          options.sizing,
        ),
      ) as Pick<AddSvgImageOptions, 'width' | 'height' | 'sourceRectangle'>;
  if (resolved.info.contentType !== 'image/svg+xml') {
    return slide.addImage(resolved.bytes, {
      ...options.imageOptions,
      ...(placement ?? {}),
      contentType: resolved.info.contentType,
    });
  }
  return slide.addSvgImage(resolved.bytes, prepared.fallbackPngBytes!, {
    ...options.imageOptions,
    ...(placement ?? {}),
  });
}

function imageSizingForPlaceholder(
  slide: SlideModel,
  selector: PlaceholderSelector | undefined,
  sizing: Readonly<ImageSizing>,
): Readonly<ImageSizing> {
  if (selector === undefined) return sizing;
  const owner = slide.placeholders.find((shape) => {
    const identity = shape.placeholder;
    return typeof selector === 'string'
      ? shape.name === selector
      : identity?.type === selector.type && identity.index === selector.index;
  });
  if (owner?.placeholder?.type !== 'pic') return sizing;
  return Object.freeze({
    ...sizing,
    width: owner.transform.width,
    height: owner.transform.height,
  });
}
