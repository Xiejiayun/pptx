import { type LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import { replaceSlideNumberCachedText } from '@pptx/codecs';
import type { PresentationModel } from './presentation.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const MIN_INT32 = -2_147_483_648;
const MAX_INT32 = 2_147_483_647;
const INVALID = Symbol('invalid first slide number');

export function normalizeFirstSlideNumber(value: unknown): number {
  if (typeof value !== 'number' || !Number.isSafeInteger(value)) {
    throw new TypeError('First slide number must be a safe integer');
  }
  if (value < MIN_INT32 || value > MAX_INT32) {
    throw new RangeError('First slide number must fit the OOXML signed Int32 range');
  }
  return value === 0 ? 0 : value;
}

export function readFirstSlideNumber(xml: LosslessXmlDocument): number | undefined {
  const root = presentationRoot(xml);
  if (!root) return undefined;
  const value = directAttribute(root, 'firstSlideNum');
  if (typeof value !== 'string' || !/^[+-]?\d+$/.test(value)) return undefined;
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) && parsed >= MIN_INT32 && parsed <= MAX_INT32
    ? parsed === 0 ? 0 : parsed
    : undefined;
}

export function replaceFirstSlideNumber(
  xml: LosslessXmlDocument,
  value: number | undefined,
): boolean {
  const normalized = value === undefined ? undefined : normalizeFirstSlideNumber(value);
  const root = presentationRoot(xml);
  if (!root) throw new Error('Presentation root is missing or ambiguous');
  const matches = root.attributes.filter((attribute) => attribute.name === 'firstSlideNum');
  if (matches.length > 1) {
    throw new Error('Presentation first slide number is ambiguous');
  }
  const attribute = matches[0];
  if (normalized === undefined) {
    if (!attribute) return false;
    let start = attribute.start;
    while (start > root.start && /[\t ]/.test(xml.source[start - 1] ?? '')) start -= 1;
    xml.replace(start, attribute.end, '');
    return true;
  }
  if (attribute) {
    const current = /^[+-]?\d+$/.test(attribute.value)
      ? Number(attribute.value)
      : undefined;
    if (current === normalized) return false;
    xml.replaceAttribute(attribute, String(normalized));
    return true;
  }
  const marker = root.selfClosing
    ? xml.source.lastIndexOf('/', root.startTagEnd - 1)
    : root.startTagEnd - 1;
  if (marker < root.start) throw new Error('Presentation root start tag is invalid');
  xml.replace(marker, marker, ` firstSlideNum="${normalized}"`);
  return true;
}

export function synchronizeSlideNumberCaches(model: PresentationModel): void {
  const first = model.firstSlideNumber ?? 1;
  const slides = model.slides;
  for (const [index, slide] of slides.entries()) {
    const effective = first + index;
    if (!Number.isSafeInteger(effective)) {
      throw new RangeError('Effective slide number exceeds the JavaScript safe integer range');
    }
    replaceSlideNumberCachedText(model.opcPackage, slide.partUri, String(effective));
  }
}

function presentationRoot(xml: LosslessXmlDocument): XmlElement | undefined {
  return xml.roots.length === 1
    && xml.roots[0]?.localName === 'presentation'
    && elementNamespaceUri(xml.roots[0]) === PRESENTATION_NAMESPACE
    ? xml.roots[0]
    : undefined;
}

function directAttribute(
  element: XmlElement,
  name: string,
): string | undefined | typeof INVALID {
  const matches = element.attributes.filter((attribute) => attribute.name === name);
  if (matches.length > 1) return INVALID;
  return matches[0]?.value;
}

function elementNamespaceUri(element: XmlElement): string | undefined {
  const separator = element.name.indexOf(':');
  const prefix = separator < 0 ? '' : element.name.slice(0, separator);
  const declarationName = prefix.length === 0 ? 'xmlns' : `xmlns:${prefix}`;
  let current: XmlElement | undefined = element;
  while (current) {
    const declarations = current.attributes.filter(
      (attribute) => attribute.name === declarationName,
    );
    if (declarations.length > 1) return undefined;
    if (declarations[0]) return declarations[0].value;
    current = current.parent;
  }
  return undefined;
}
