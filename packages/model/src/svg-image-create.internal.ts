import { escapeXmlAttribute } from '@pptx/lossless-xml';
import {
  normalizeEmbeddedRasterImage,
  renderEmbeddedImageXml,
  type NormalizedEmbeddedImageAppearance,
} from './image-create.internal.js';
import type { PlaceholderIdentity } from './placeholder.js';
import type { SlideSize } from './units.js';

export const SVG_IMAGE_EXTENSION_URI =
  '{96DAC541-7B7A-43D3-8B79-37D633B846F1}';
export const SVG_IMAGE_NAMESPACE =
  'http://schemas.microsoft.com/office/drawing/2016/SVG/main';

const OPTION_KEYS = new Set([
  'name',
  'altText',
  'placeholder',
  'x',
  'y',
  'width',
  'height',
  'rotation',
  'flipHorizontal',
  'flipVertical',
  'sourceRectangle',
]);

export interface NormalizedEmbeddedSvgImage
  extends NormalizedEmbeddedImageAppearance {
  readonly svgBytes: Uint8Array;
  readonly fallbackPngBytes: Uint8Array;
}

export function normalizeEmbeddedSvgImage(
  svgBytes: unknown,
  fallbackPngBytes: unknown,
  options: unknown = {},
  slideSize?: Readonly<SlideSize>,
): NormalizedEmbeddedSvgImage {
  const detachedSvgBytes = normalizeSvgBytes(svgBytes);
  const rasterOptions = normalizeSvgOptions(options);
  const fallback = normalizeEmbeddedRasterImage(fallbackPngBytes, rasterOptions, slideSize);
  const {
    bytes: detachedFallbackPngBytes,
    contentType: _contentType,
    extension: _extension,
    ...appearance
  } = fallback;
  return Object.freeze({
    svgBytes: detachedSvgBytes,
    fallbackPngBytes: detachedFallbackPngBytes,
    ...appearance,
  });
}

export function renderEmbeddedSvgImageXml(
  id: number,
  definition: Readonly<NormalizedEmbeddedSvgImage>,
  fallbackRelationshipId: string,
  svgRelationshipId: string,
  defaultName: string,
  placeholder?: Readonly<PlaceholderIdentity>,
): string {
  const fallbackEmbed = escapeXmlAttribute(fallbackRelationshipId);
  const svgEmbed = escapeXmlAttribute(svgRelationshipId);
  const blip = `<a:blip r:embed="${fallbackEmbed}"><a:extLst>`
    + `<a:ext uri="${SVG_IMAGE_EXTENSION_URI}">`
    + `<asvg:svgBlip xmlns:asvg="${SVG_IMAGE_NAMESPACE}" r:embed="${svgEmbed}"/>`
    + '</a:ext></a:extLst></a:blip>';
  return renderEmbeddedImageXml(id, definition, blip, defaultName, placeholder);
}

function normalizeSvgBytes(value: unknown): Uint8Array {
  if (!(value instanceof Uint8Array)) {
    throw new TypeError('Embedded SVG image bytes must be a Uint8Array');
  }
  if (value.length === 0) {
    throw new RangeError('Embedded SVG image bytes must not be empty');
  }
  return new Uint8Array(value);
}

function normalizeSvgOptions(value: unknown): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError('Embedded SVG image options must be an object');
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError('Embedded SVG image options must be an ordinary object');
  }
  const result = Object.create(null) as Record<string, unknown>;
  result.contentType = 'image/png';
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !OPTION_KEYS.has(key)) {
      throw new TypeError(
        `Embedded SVG image options contain unsupported property ${String(key)}`,
      );
    }
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`Embedded SVG image option ${key} must be a data property`);
    }
    result[key] = descriptor.value;
  }
  if (result.altText === undefined) result.altText = 'preencoded.svg';
  return result;
}
