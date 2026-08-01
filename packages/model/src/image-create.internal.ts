import { escapeXmlAttribute } from '@pptx/lossless-xml';
import type { RasterImageContentType } from './image.js';
import {
  normalizeImageSourceRectangle,
  renderImageSourceRectangle,
  type NormalizedImageSourceRectangle,
} from './image-source-rectangle.internal.js';
import { EMU_PER_INCH } from './units.js';

const MAX_ROTATION = 21_600_000;
const OPTION_KEYS = new Set([
  'contentType',
  'name',
  'altText',
  'x',
  'y',
  'width',
  'height',
  'rotation',
  'flipHorizontal',
  'flipVertical',
  'sourceRectangle',
]);

export interface NormalizedEmbeddedImageAppearance {
  readonly name: string | undefined;
  readonly altText: string;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly rotation: number;
  readonly flipHorizontal: boolean;
  readonly flipVertical: boolean;
  readonly sourceRectangle?: Readonly<NormalizedImageSourceRectangle>;
}

export interface NormalizedEmbeddedRasterImage
  extends NormalizedEmbeddedImageAppearance {
  readonly bytes: Uint8Array;
  readonly contentType: RasterImageContentType;
  readonly extension: '.png' | '.jpeg' | '.gif';
}

export function normalizeEmbeddedRasterImage(
  bytes: unknown,
  options: unknown,
): NormalizedEmbeddedRasterImage {
  if (!(bytes instanceof Uint8Array)) {
    throw new TypeError('Embedded raster image bytes must be a Uint8Array');
  }
  if (bytes.length === 0) {
    throw new RangeError('Embedded raster image bytes must not be empty');
  }

  const values = readOptions(options);
  const contentType = normalizeContentType(values.contentType);
  const name = normalizeXmlString(values.name, undefined, 'name');
  const altText = normalizeXmlString(values.altText, 'preencoded.png', 'altText');
  const width = normalizeInteger(values.width, EMU_PER_INCH, 'width');
  const height = normalizeInteger(values.height, EMU_PER_INCH, 'height');
  if (width <= 0) throw new RangeError('Embedded raster image width must be positive');
  if (height <= 0) throw new RangeError('Embedded raster image height must be positive');
  const rotation = normalizeInteger(values.rotation, 0, 'rotation');
  if (rotation < -MAX_ROTATION || rotation > MAX_ROTATION) {
    throw new RangeError(
      'Embedded raster image rotation must be between -21600000 and 21600000',
    );
  }
  const sourceRectangle = values.sourceRectangle === undefined
    ? undefined
    : normalizeImageSourceRectangle(
        values.sourceRectangle,
        'Embedded raster image sourceRectangle',
      );

  return Object.freeze({
    bytes: new Uint8Array(bytes),
    contentType,
    extension: extensionFor(contentType),
    name,
    altText,
    x: normalizeInteger(values.x, 0, 'x'),
    y: normalizeInteger(values.y, 0, 'y'),
    width,
    height,
    rotation,
    flipHorizontal: normalizeBoolean(values.flipHorizontal, false, 'flipHorizontal'),
    flipVertical: normalizeBoolean(values.flipVertical, false, 'flipVertical'),
    ...(sourceRectangle === undefined ? {} : { sourceRectangle }),
  });
}

export function renderEmbeddedRasterImageXml(
  id: number,
  definition: Readonly<NormalizedEmbeddedRasterImage>,
  relationshipId: string,
  defaultName: string,
): string {
  const embed = escapeXmlAttribute(relationshipId);
  return renderEmbeddedImageXml(
    id,
    definition,
    `<a:blip r:embed="${embed}"/>`,
    defaultName,
  );
}

export function renderEmbeddedImageXml(
  id: number,
  definition: Readonly<NormalizedEmbeddedImageAppearance>,
  blipXml: string,
  defaultName: string,
): string {
  const name = escapeXmlAttribute(definition.name ?? defaultName);
  const altText = escapeXmlAttribute(definition.altText);
  const transformAttributes = [
    definition.rotation === 0 ? '' : ` rot="${definition.rotation}"`,
    definition.flipHorizontal ? ' flipH="1"' : '',
    definition.flipVertical ? ' flipV="1"' : '',
  ].join('');
  const sourceRectangle = definition.sourceRectangle
    ? renderImageSourceRectangle(definition.sourceRectangle)
    : '';

  return '<p:pic>'
    + `<p:nvPicPr><p:cNvPr id="${id}" name="${name}" descr="${altText}"/>`
    + '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>'
    + `<p:blipFill>${blipXml}${sourceRectangle}`
    + '<a:stretch><a:fillRect/></a:stretch></p:blipFill>'
    + `<p:spPr><a:xfrm${transformAttributes}>`
    + `<a:off x="${definition.x}" y="${definition.y}"/>`
    + `<a:ext cx="${definition.width}" cy="${definition.height}"/>`
    + '</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
    + '</p:pic>';
}

function readOptions(value: unknown): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError('Embedded raster image options must be an object');
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError('Embedded raster image options must be an ordinary object');
  }
  const result = Object.create(null) as Record<string, unknown>;
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !OPTION_KEYS.has(key)) {
      throw new TypeError(
        `Embedded raster image options contain unsupported property ${String(key)}`,
      );
    }
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`Embedded raster image option ${key} must be a data property`);
    }
    result[key] = descriptor.value;
  }
  return result;
}

function normalizeContentType(value: unknown): RasterImageContentType {
  if (value === 'image/png' || value === 'image/jpeg' || value === 'image/gif') return value;
  throw new TypeError('Embedded raster image contentType is unsupported');
}

function extensionFor(
  contentType: RasterImageContentType,
): NormalizedEmbeddedRasterImage['extension'] {
  if (contentType === 'image/png') return '.png';
  if (contentType === 'image/jpeg') return '.jpeg';
  return '.gif';
}

function normalizeXmlString<T extends string | undefined>(
  value: unknown,
  defaultValue: T,
  name: string,
): string | T {
  if (value === undefined) return defaultValue;
  if (typeof value !== 'string') {
    throw new TypeError(`Embedded raster image ${name} must be a string`);
  }
  if (!isValidXmlString(value)) {
    throw new TypeError(`Embedded raster image ${name} contains invalid XML characters`);
  }
  return value;
}

function normalizeInteger(value: unknown, defaultValue: number, name: string): number {
  if (value === undefined) return defaultValue;
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`Embedded raster image ${name} must be finite`);
  }
  if (!Number.isSafeInteger(value)) {
    throw new RangeError(`Embedded raster image ${name} must be a safe integer`);
  }
  return value === 0 ? 0 : value;
}

function normalizeBoolean(value: unknown, defaultValue: boolean, name: string): boolean {
  if (value === undefined) return defaultValue;
  if (typeof value !== 'boolean') {
    throw new TypeError(`Embedded raster image ${name} must be a boolean`);
  }
  return value;
}

function isValidXmlString(value: string): boolean {
  for (const character of value) {
    const codePoint = character.codePointAt(0)!;
    if (
      codePoint === 0x09
      || codePoint === 0x0A
      || codePoint === 0x0D
      || (codePoint >= 0x20 && codePoint <= 0xD7FF)
      || (codePoint >= 0xE000 && codePoint <= 0xFFFD)
      || (codePoint >= 0x10000 && codePoint <= 0x10FFFF)
    ) continue;
    return false;
  }
  return true;
}
