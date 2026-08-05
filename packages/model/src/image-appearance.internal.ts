import {
  escapeXmlAttribute,
  type LosslessXmlDocument,
  type XmlAttribute,
  type XmlElement,
} from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const IMAGE_OWNER_NAMES = Object.freeze({
  sp: 'nvSpPr',
  pic: 'nvPicPr',
  graphicFrame: 'nvGraphicFramePr',
  grpSp: 'nvGrpSpPr',
} as const);
const BLIP_EFFECT_NAMES = new Set([
  'alphaBiLevel',
  'alphaCeiling',
  'alphaFloor',
  'alphaInv',
  'alphaMod',
  'alphaModFix',
  'alphaRepl',
  'biLevel',
  'blur',
  'clrChange',
  'clrRepl',
  'duotone',
  'fillOverlay',
  'grayscl',
  'hsl',
  'lum',
  'tint',
  'extLst',
]);

export function normalizeImageName(value: unknown): string {
  return normalizeXmlString(value, 'Image name', false)!;
}

export function normalizeImageAltText(value: unknown): string | undefined {
  return normalizeXmlString(value, 'Image alt text', true);
}

export function normalizeImageRounding(value: unknown): boolean {
  if (typeof value !== 'boolean') throw new TypeError('Image rounding must be a boolean');
  return value;
}

export function normalizeImageTransparency(value: unknown): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError('Image transparency must be finite');
  }
  if (value < 0 || value > 100) {
    throw new RangeError('Image transparency must be between 0 and 100');
  }
  const rounded = Math.round(value * 1000) / 1000;
  return rounded === 0 ? 0 : rounded;
}

export function readImageAltText(
  _xml: LosslessXmlDocument,
  picture: XmlElement,
  _partUri: string,
): string | undefined {
  if (picture.localName !== 'pic' || namespaceUri(picture) !== PRESENTATION_NAMESPACE) {
    return undefined;
  }
  const properties = inspectNonVisualProperties(picture);
  if (!properties) return undefined;
  const attributes = properties.attributes.filter(({ name }) => name === 'descr');
  return attributes.length === 1 ? attributes[0]!.value : undefined;
}

export function replaceShapeName(
  xml: LosslessXmlDocument,
  shape: XmlElement,
  name: string,
  partUri: string,
): boolean {
  return replaceMetadataAttribute(
    xml,
    resolveNonVisualProperties(shape, partUri),
    'name',
    name,
    partUri,
  );
}

export function replaceImageAltText(
  xml: LosslessXmlDocument,
  picture: XmlElement,
  altText: string | undefined,
  partUri: string,
): boolean {
  if (picture.localName !== 'pic' || namespaceUri(picture) !== PRESENTATION_NAMESPACE) {
    throw new ModelParseError('Image alt text is not safely editable', partUri);
  }
  return replaceMetadataAttribute(
    xml,
    resolveNonVisualProperties(picture, partUri),
    'descr',
    altText,
    partUri,
  );
}

export function readImageRounding(picture: XmlElement): boolean | undefined {
  const state = inspectImageGeometry(picture);
  return state?.type === 'ellipse' ? true : state?.type === 'rect' ? false : undefined;
}

export function replaceImageRounding(
  xml: LosslessXmlDocument,
  picture: XmlElement,
  rounding: boolean,
  partUri: string,
): boolean {
  const state = inspectImageGeometry(picture);
  if (!state) throw new ModelParseError('Image rounding is not safely editable', partUri);
  const type = rounding ? 'ellipse' : 'rect';
  if (state.type === type) return false;
  xml.replaceAttribute(state.typeAttribute, type);
  return true;
}

export function readImageTransparency(picture: XmlElement): number | undefined {
  return inspectImageTransparency(picture)?.transparency;
}

export function replaceImageTransparency(
  xml: LosslessXmlDocument,
  picture: XmlElement,
  transparency: number,
  partUri: string,
): boolean {
  const state = inspectImageTransparency(picture);
  if (!state) throw new ModelParseError('Image transparency is not safely editable', partUri);
  if (transparency === 0) {
    if (!state.alpha) return false;
    xml.removeElement(state.alpha.element);
    return true;
  }
  if (state.transparency === transparency) return false;

  const amount = String(Math.round((100 - transparency) * 1000));
  if (state.alpha) {
    xml.replaceAttribute(state.alpha.amount, amount);
    return true;
  }
  const prefix = lexicalPrefix(state.blip.name);
  const qualified = prefix === '' ? '' : `${prefix}:`;
  const encoded = `<${qualified}alphaModFix amt="${amount}"/>`;
  if (state.blip.selfClosing) {
    const source = xml.original(state.blip);
    const marker = source.lastIndexOf('/>');
    xml.replaceElement(
      state.blip,
      source.slice(0, marker) + `>${encoded}</${state.blip.name}>`,
    );
    return true;
  }
  const position = state.extensionList?.start ?? state.blip.endTagStart;
  xml.replace(position, position, encoded);
  return true;
}

function resolveNonVisualProperties(shape: XmlElement, partUri: string): XmlElement {
  const properties = inspectNonVisualProperties(shape);
  if (!properties) {
    throw new ModelParseError('Shape identity is not safely editable', partUri);
  }
  return properties;
}

function inspectNonVisualProperties(shape: XmlElement): XmlElement | undefined {
  if (namespaceUri(shape) !== PRESENTATION_NAMESPACE) {
    return undefined;
  }
  const ownerName = IMAGE_OWNER_NAMES[shape.localName as keyof typeof IMAGE_OWNER_NAMES];
  if (!ownerName) return undefined;
  const owners = directChildren(shape).filter((child) => child.localName === ownerName);
  const owner = owners.length === 1 && namespaceUri(owners[0]!) === PRESENTATION_NAMESPACE
    ? owners[0]
    : undefined;
  if (!owner) return undefined;
  const properties = directChildren(owner).filter((child) => child.localName === 'cNvPr');
  return properties.length === 1 && namespaceUri(properties[0]!) === PRESENTATION_NAMESPACE
    ? properties[0]
    : undefined;
}

function metadataAttribute(
  _xml: LosslessXmlDocument,
  properties: XmlElement,
  name: 'name' | 'descr',
  partUri: string,
): XmlAttribute | undefined {
  const candidates = properties.attributes.filter((attribute) => attribute.name === name);
  if (candidates.length > 1) {
    throw new ModelParseError(`Shape has repeated ${name} attributes`, partUri);
  }
  return candidates[0];
}

function replaceMetadataAttribute(
  xml: LosslessXmlDocument,
  properties: XmlElement,
  name: 'name' | 'descr',
  value: string | undefined,
  partUri: string,
): boolean {
  const current = metadataAttribute(xml, properties, name, partUri);
  if (value === undefined) {
    if (!current) return false;
    removeAttribute(xml, properties, current);
    return true;
  }
  if (current?.value === value) return false;
  if (current) xml.replaceAttribute(current, value);
  else insertAttribute(xml, properties, name, value);
  return true;
}

function inspectImageGeometry(
  picture: XmlElement,
): { readonly type: 'rect' | 'ellipse'; readonly typeAttribute: XmlAttribute } | undefined {
  if (picture.localName !== 'pic' || namespaceUri(picture) !== PRESENTATION_NAMESPACE) {
    return undefined;
  }
  const properties = directPictureChild(picture, 'spPr', PRESENTATION_NAMESPACE);
  if (!properties) return undefined;
  const geometries = directChildren(properties).filter(
    ({ localName }) => localName === 'prstGeom' || localName === 'custGeom',
  );
  const geometry = geometries.length === 1 ? geometries[0] : undefined;
  if (!geometry || geometry.localName !== 'prstGeom' || namespaceUri(geometry) !== DRAWING_NAMESPACE) {
    return undefined;
  }
  const attributes = nonNamespaceAttributes(geometry).filter(({ localName }) => localName === 'prst');
  const typeAttribute = attributes.length === 1 && attributes[0]!.name === 'prst'
    ? attributes[0]
    : undefined;
  if (!typeAttribute || (typeAttribute.value !== 'rect' && typeAttribute.value !== 'ellipse')) {
    return undefined;
  }
  return { type: typeAttribute.value, typeAttribute };
}

function inspectImageTransparency(picture: XmlElement): {
  readonly blip: XmlElement;
  readonly extensionList?: XmlElement;
  readonly alpha?: { readonly element: XmlElement; readonly amount: XmlAttribute };
  readonly transparency: number;
} | undefined {
  if (picture.localName !== 'pic' || namespaceUri(picture) !== PRESENTATION_NAMESPACE) {
    return undefined;
  }
  const blipFill = directPictureChild(picture, 'blipFill', PRESENTATION_NAMESPACE);
  const blip = blipFill ? directPictureChild(blipFill, 'blip', DRAWING_NAMESPACE) : undefined;
  if (!blip || hasNonWhitespaceText(blip)) return undefined;
  const children = directChildren(blip);
  if (children.some((child) =>
    namespaceUri(child) !== DRAWING_NAMESPACE || !BLIP_EFFECT_NAMES.has(child.localName))) {
    return undefined;
  }
  const extensionLists = children.filter(({ localName }) => localName === 'extLst');
  if (
    extensionLists.length > 1
    || (extensionLists[0] !== undefined && children.at(-1) !== extensionLists[0])
  ) return undefined;
  const alphas = children.filter(({ localName }) => localName === 'alphaModFix');
  if (alphas.length > 1) return undefined;
  const alpha = alphas[0];
  if (!alpha) {
    return { blip, ...(extensionLists[0] ? { extensionList: extensionLists[0] } : {}), transparency: 0 };
  }
  if (hasNonWhitespaceText(alpha) || directChildren(alpha).length > 0) return undefined;
  const attributes = nonNamespaceAttributes(alpha);
  const amount = attributes.length === 1 && attributes[0]!.name === 'amt'
    ? attributes[0]
    : undefined;
  if (!amount || !/^(?:0|[1-9]\d*)$/u.test(amount.value)) return undefined;
  const numeric = Number(amount.value);
  if (!Number.isSafeInteger(numeric) || numeric < 0 || numeric > 100_000) return undefined;
  return {
    blip,
    ...(extensionLists[0] ? { extensionList: extensionLists[0] } : {}),
    alpha: { element: alpha, amount },
    transparency: (100_000 - numeric) / 1000,
  };
}

function directPictureChild(
  parent: XmlElement,
  localName: string,
  namespace: string,
): XmlElement | undefined {
  const candidates = directChildren(parent).filter((child) => child.localName === localName);
  return candidates.length === 1 && namespaceUri(candidates[0]!) === namespace
    ? candidates[0]
    : undefined;
}

function insertAttribute(
  xml: LosslessXmlDocument,
  element: XmlElement,
  name: string,
  value: string,
): void {
  const position = element.selfClosing
    ? xml.source.lastIndexOf('/', element.startTagEnd - 1)
    : element.startTagEnd - 1;
  if (position <= element.start) throw new Error(`Image ${element.localName} start tag is invalid`);
  xml.replace(position, position, ` ${name}="${escapeXmlAttribute(value)}"`);
}

function removeAttribute(
  xml: LosslessXmlDocument,
  element: XmlElement,
  attribute: XmlAttribute,
): void {
  let start = attribute.start;
  while (start > element.start && /[\t ]/u.test(xml.source[start - 1] ?? '')) start -= 1;
  xml.replace(start, attribute.end, '');
}

function normalizeXmlString(
  value: unknown,
  context: string,
  allowUndefined: boolean,
): string | undefined {
  if (value === undefined && allowUndefined) return undefined;
  if (typeof value !== 'string') throw new TypeError(`${context} must be a string`);
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
    throw new TypeError(`${context} contains invalid XML characters`);
  }
  return value;
}

function directChildren(element: XmlElement): XmlElement[] {
  return element.children.filter((child): child is XmlElement => child.type === 'element');
}

function nonNamespaceAttributes(element: XmlElement): XmlAttribute[] {
  return element.attributes.filter(({ name }) => name !== 'xmlns' && !name.startsWith('xmlns:'));
}

function hasNonWhitespaceText(element: XmlElement): boolean {
  return element.children.some((child) => child.type === 'text' && /\S/u.test(child.value));
}

function namespaceUri(element: XmlElement): string | undefined {
  const prefix = lexicalPrefix(element.name);
  const declarationName = prefix === '' ? 'xmlns' : `xmlns:${prefix}`;
  for (let current: XmlElement | undefined = element; current; current = current.parent) {
    const declarations = current.attributes.filter(({ name }) => name === declarationName);
    if (declarations.length > 1) return undefined;
    if (declarations[0]) return declarations[0].value;
  }
  return undefined;
}

function lexicalPrefix(name: string): string {
  const separator = name.indexOf(':');
  return separator < 0 ? '' : name.slice(0, separator);
}
