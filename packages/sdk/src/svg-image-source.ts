import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import type { SvgImageContentType } from '@pptx/model';

const SVG_NAMESPACE = 'http://www.w3.org/2000/svg';
const NUMBER_PATTERN = '[+-]?(?:\\d+(?:\\.\\d*)?|\\.\\d+)(?:[eE][+-]?\\d+)?';
const DIMENSION_PATTERN = new RegExp(`^\\s*(${NUMBER_PATTERN})([A-Za-z]*)\\s*$`);
const VIEW_BOX_NUMBER_PATTERN = new RegExp(NUMBER_PATTERN, 'g');

export interface SvgImageInfo {
  readonly contentType: SvgImageContentType;
  readonly width: number;
  readonly height: number;
}

export function inspectSvgImage(bytes: Uint8Array): SvgImageInfo {
  if (!(bytes instanceof Uint8Array)) {
    throw new TypeError('SVG image bytes must be a Uint8Array');
  }
  if (bytes.byteLength === 0) throw new TypeError('SVG image bytes cannot be empty');

  let source: string;
  try {
    source = new TextDecoder('utf-8', { fatal: true }).decode(bytes);
  } catch {
    throw new TypeError('SVG image bytes must contain valid UTF-8');
  }

  let xml: LosslessXmlDocument;
  try {
    xml = LosslessXmlDocument.parse(source);
  } catch {
    throw new TypeError('Invalid SVG XML');
  }
  if (xml.roots.length !== 1) throw new TypeError('SVG XML must contain exactly one root element');
  const root = xml.roots[0]!;
  if (
    root.localName !== 'svg'
    || elementNamespaceUri(root) !== SVG_NAMESPACE
    || !containsOnlyXmlMisc(source.slice(0, root.start))
    || !containsOnlyXmlMisc(source.slice(root.end))
  ) {
    throw new TypeError('SVG XML root element is invalid');
  }

  const viewBoxValue = uniqueAttribute(root, 'viewBox');
  const viewBox = viewBoxValue === undefined ? undefined : parseViewBox(viewBoxValue);
  const width = parseDimension(uniqueAttribute(root, 'width'));
  const height = parseDimension(uniqueAttribute(root, 'height'));
  let resolvedWidth: number;
  let resolvedHeight: number;
  if (width !== undefined && height !== undefined) {
    resolvedWidth = width;
    resolvedHeight = height;
  } else if (viewBox && width !== undefined) {
    resolvedWidth = width;
    resolvedHeight = width * viewBox.height / viewBox.width;
  } else if (viewBox && height !== undefined) {
    resolvedWidth = height * viewBox.width / viewBox.height;
    resolvedHeight = height;
  } else if (viewBox) {
    resolvedWidth = viewBox.width;
    resolvedHeight = viewBox.height;
  } else {
    resolvedWidth = 300;
    resolvedHeight = 150;
  }
  if (!isPositiveFinite(resolvedWidth) || !isPositiveFinite(resolvedHeight)) {
    throw new TypeError('SVG intrinsic dimensions are invalid');
  }
  return Object.freeze({
    contentType: 'image/svg+xml',
    width: resolvedWidth,
    height: resolvedHeight,
  });
}

function elementNamespaceUri(element: XmlElement): string | undefined {
  const separator = element.name.indexOf(':');
  const prefix = separator < 0 ? '' : element.name.slice(0, separator);
  const name = prefix.length === 0 ? 'xmlns' : `xmlns:${prefix}`;
  const declarations = element.attributes.filter((attribute) => attribute.name === name);
  return declarations.length === 1 ? declarations[0]!.value : undefined;
}

function uniqueAttribute(element: XmlElement, name: string): string | undefined {
  const attributes = element.attributes.filter((attribute) => attribute.name === name);
  if (attributes.length > 1) throw new TypeError(`SVG ${name} attribute is duplicated`);
  return attributes[0]?.value;
}

function parseViewBox(value: string): { readonly width: number; readonly height: number } {
  const matches = [...value.matchAll(VIEW_BOX_NUMBER_PATTERN)];
  if (matches.length !== 4) throw new TypeError('SVG viewBox is invalid');
  let cursor = 0;
  const values: number[] = [];
  for (const match of matches) {
    const separator = value.slice(cursor, match.index);
    if (
      (cursor === 0 && !/^\s*$/.test(separator))
      || (cursor > 0 && !/^(?:\s+|\s*,\s*)$/.test(separator))
    ) throw new TypeError('SVG viewBox is invalid');
    values.push(Number(match[0]));
    cursor = match.index! + match[0].length;
  }
  if (!/^\s*$/.test(value.slice(cursor))) throw new TypeError('SVG viewBox is invalid');
  const [x, y, width, height] = values as [number, number, number, number];
  if (
    !Number.isFinite(x)
    || !Number.isFinite(y)
    || !isPositiveFinite(width)
    || !isPositiveFinite(height)
  ) throw new TypeError('SVG viewBox dimensions must be positive and finite');
  return Object.freeze({ width, height });
}

function parseDimension(value: string | undefined): number | undefined {
  if (value === undefined) return undefined;
  const match = DIMENSION_PATTERN.exec(value);
  if (!match) return undefined;
  const amount = Number(match[1]);
  if (!isPositiveFinite(amount)) return undefined;
  const scale = dimensionUnitScale(match[2]!.toLowerCase());
  if (scale === undefined) return undefined;
  const pixels = amount * scale;
  return isPositiveFinite(pixels) ? pixels : undefined;
}

function dimensionUnitScale(unit: string): number | undefined {
  switch (unit) {
    case '':
    case 'px': return 1;
    case 'in': return 96;
    case 'cm': return 96 / 2.54;
    case 'mm': return 96 / 25.4;
    case 'pt': return 96 / 72;
    case 'pc': return 16;
    case 'q': return 96 / 101.6;
    default: return undefined;
  }
}

function containsOnlyXmlMisc(value: string): boolean {
  let cursor = 0;
  while (cursor < value.length) {
    const whitespace = /^\s+/.exec(value.slice(cursor));
    if (whitespace) {
      cursor += whitespace[0].length;
      continue;
    }
    if (value.startsWith('<!--', cursor)) {
      const end = value.indexOf('-->', cursor + 4);
      if (end < 0) return false;
      cursor = end + 3;
      continue;
    }
    if (value.startsWith('<?', cursor)) {
      const end = value.indexOf('?>', cursor + 2);
      if (end < 0) return false;
      cursor = end + 2;
      continue;
    }
    return false;
  }
  return true;
}

function isPositiveFinite(value: number): boolean {
  return Number.isFinite(value) && value > 0;
}
