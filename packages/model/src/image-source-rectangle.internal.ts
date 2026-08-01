import type {
  LosslessXmlDocument,
  XmlAttribute,
  XmlElement,
} from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const EDGE_KEYS = ['left', 'top', 'right', 'bottom'] as const;
const EDGE_KEY_SET = new Set<string>(EDGE_KEYS);
const ATTRIBUTE_TO_EDGE: ReadonlyMap<string, (typeof EDGE_KEYS)[number]> = new Map([
  ['l', 'left'],
  ['t', 'top'],
  ['r', 'right'],
  ['b', 'bottom'],
] as const);
const PERCENT_SCALE = 1_000;
const FULL_PERCENT = 100 * PERCENT_SCALE;
const INT32_MIN = -2_147_483_648;
const INT32_MAX = 2_147_483_647;

export interface NormalizedImageSourceRectangle {
  readonly left: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
}

export function normalizeImageSourceRectangle(
  value: unknown,
  context: string,
): Readonly<NormalizedImageSourceRectangle> {
  const values = readRectangle(value, context);
  const raw = {
    left: normalizeEdge(values.left, `${context} left`),
    top: normalizeEdge(values.top, `${context} top`),
    right: normalizeEdge(values.right, `${context} right`),
    bottom: normalizeEdge(values.bottom, `${context} bottom`),
  };
  if (raw.left + raw.right >= FULL_PERCENT) {
    throw new RangeError(`${context} left and right must leave a positive source width`);
  }
  if (raw.top + raw.bottom >= FULL_PERCENT) {
    throw new RangeError(`${context} top and bottom must leave a positive source height`);
  }
  return Object.freeze({
    left: percentage(raw.left),
    top: percentage(raw.top),
    right: percentage(raw.right),
    bottom: percentage(raw.bottom),
  });
}

export function renderImageSourceRectangle(
  value: Readonly<NormalizedImageSourceRectangle>,
  prefix = 'a',
): string {
  const name = prefix.length === 0 ? 'srcRect' : `${prefix}:srcRect`;
  return `<${name} l="${rawPercentage(value.left)}" t="${rawPercentage(value.top)}" `
    + `r="${rawPercentage(value.right)}" b="${rawPercentage(value.bottom)}"/>`;
}

export function readImageSourceRectangle(
  _xml: LosslessXmlDocument,
  picture: XmlElement,
): Readonly<NormalizedImageSourceRectangle> | undefined {
  return inspectOwner(picture)?.sourceRectangle?.snapshot;
}

export function replaceImageSourceRectangle(
  xml: LosslessXmlDocument,
  picture: XmlElement,
  value: Readonly<NormalizedImageSourceRectangle> | undefined,
  partUri: string,
): boolean {
  const owner = inspectOwner(picture);
  if (!owner) {
    throw new ModelParseError('Image source rectangle state is not safely editable', partUri);
  }
  const current = owner.sourceRectangle;
  if (current?.snapshot && rectanglesEqual(current.snapshot, value)) return false;
  if (!current) {
    if (!value) return false;
    const rendered = renderForParent(value, owner.prefix, owner.fill);
    const fillChoice = owner.children.find(
      (child) =>
        (child.localName === 'tile' || child.localName === 'stretch')
        && namespaceUri(child) === DRAWING_NAMESPACE,
    );
    xml.replace(fillChoice?.start ?? owner.blip.end, fillChoice?.start ?? owner.blip.end, rendered);
    return true;
  }
  if (!value) {
    xml.removeElement(current.element);
    return true;
  }
  xml.replaceElement(
    current.element,
    renderForParent(value, lexicalPrefix(current.element.name), owner.fill),
  );
  return true;
}

interface ExistingSourceRectangle {
  readonly element: XmlElement;
  readonly snapshot?: Readonly<NormalizedImageSourceRectangle>;
}

interface SourceRectangleOwner {
  readonly fill: XmlElement;
  readonly children: readonly XmlElement[];
  readonly blip: XmlElement;
  readonly sourceRectangle?: ExistingSourceRectangle;
  readonly prefix: string;
}

function inspectOwner(picture: XmlElement): SourceRectangleOwner | undefined {
  if (picture.localName !== 'pic' || namespaceUri(picture) !== PRESENTATION_NAMESPACE) {
    return undefined;
  }
  const fills = directChildren(picture).filter(({ localName }) => localName === 'blipFill');
  if (fills.length !== 1 || namespaceUri(fills[0]!) !== PRESENTATION_NAMESPACE) return undefined;
  const fill = fills[0]!;
  if (hasNonWhitespaceText(fill)) return undefined;
  const children = directChildren(fill);
  const blips = children.filter(({ localName }) => localName === 'blip');
  if (blips.length !== 1 || namespaceUri(blips[0]!) !== DRAWING_NAMESPACE) return undefined;
  const rectangles = children.filter(({ localName }) => localName === 'srcRect');
  if (
    rectangles.length > 1
    || (rectangles[0] && namespaceUri(rectangles[0]) !== DRAWING_NAMESPACE)
  ) return undefined;
  const element = rectangles[0];
  const snapshot = element ? inspectSourceRectangle(element) : undefined;
  return {
    fill,
    children,
    blip: blips[0]!,
    ...(element
      ? {
          sourceRectangle: {
            element,
            ...(snapshot === undefined ? {} : { snapshot }),
          },
        }
      : {}),
    prefix: element ? lexicalPrefix(element.name) : lexicalPrefix(blips[0]!.name),
  };
}

function inspectSourceRectangle(
  element: XmlElement,
): Readonly<NormalizedImageSourceRectangle> | undefined {
  if (directChildren(element).length > 0 || hasNonWhitespaceText(element)) return undefined;
  const attributes = nonNamespaceAttributes(element);
  const raw: Record<string, number> = Object.create(null) as Record<string, number>;
  for (const attribute of attributes) {
    const edge = ATTRIBUTE_TO_EDGE.get(attribute.name);
    if (!edge || attribute.name !== attribute.localName || Object.hasOwn(raw, edge)) {
      return undefined;
    }
    const value = parseCanonicalInt32(attribute);
    if (value === undefined) return undefined;
    raw[edge] = value;
  }
  try {
    return normalizeImageSourceRectangle({
      left: (raw.left ?? 0) / PERCENT_SCALE,
      top: (raw.top ?? 0) / PERCENT_SCALE,
      right: (raw.right ?? 0) / PERCENT_SCALE,
      bottom: (raw.bottom ?? 0) / PERCENT_SCALE,
    }, 'Image source rectangle');
  } catch {
    return undefined;
  }
}

function parseCanonicalInt32(attribute: XmlAttribute): number | undefined {
  if (!/^(?:0|-?[1-9][0-9]*)$/.test(attribute.value)) return undefined;
  const value = Number(attribute.value);
  return Number.isSafeInteger(value) && value >= INT32_MIN && value <= INT32_MAX
    ? value
    : undefined;
}

function renderForParent(
  value: Readonly<NormalizedImageSourceRectangle>,
  prefix: string,
  parent: XmlElement,
): string {
  const rendered = renderImageSourceRectangle(value, prefix);
  if (namespaceUriForPrefix(parent, prefix) === DRAWING_NAMESPACE) return rendered;
  const name = prefix.length === 0 ? 'srcRect' : `${prefix}:srcRect`;
  const declaration = prefix.length === 0
    ? ` xmlns="${DRAWING_NAMESPACE}"`
    : ` xmlns:${prefix}="${DRAWING_NAMESPACE}"`;
  return rendered.replace(`<${name}`, `<${name}${declaration}`);
}

function rectanglesEqual(
  left: Readonly<NormalizedImageSourceRectangle>,
  right: Readonly<NormalizedImageSourceRectangle> | undefined,
): boolean {
  return right !== undefined
    && left.left === right.left
    && left.top === right.top
    && left.right === right.right
    && left.bottom === right.bottom;
}

function readRectangle(value: unknown, context: string): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${context} must be an object`);
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError(`${context} must be an ordinary object`);
  }
  const result = Object.create(null) as Record<string, unknown>;
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !EDGE_KEY_SET.has(key)) {
      throw new TypeError(`${context} contains unsupported property ${String(key)}`);
    }
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`${context} ${key} must be a data property`);
    }
    result[key] = descriptor.value;
  }
  for (const key of EDGE_KEYS) {
    if (!Object.hasOwn(result, key)) throw new TypeError(`${context} ${key} is required`);
  }
  return result;
}

function normalizeEdge(value: unknown, context: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  const raw = Math.round(value * PERCENT_SCALE);
  if (!Number.isSafeInteger(raw) || raw < INT32_MIN || raw > INT32_MAX) {
    throw new RangeError(`${context} must fit the OOXML Int32 percentage range`);
  }
  if (raw >= FULL_PERCENT) throw new RangeError(`${context} must be less than 100 percent`);
  return raw === 0 ? 0 : raw;
}

function percentage(raw: number): number {
  const value = raw / PERCENT_SCALE;
  return value === 0 ? 0 : value;
}

function rawPercentage(value: number): number {
  const raw = Math.round(value * PERCENT_SCALE);
  return raw === 0 ? 0 : raw;
}

function namespaceUri(element: XmlElement): string | undefined {
  return namespaceUriForPrefix(element, lexicalPrefix(element.name));
}

function namespaceUriForPrefix(
  element: XmlElement,
  prefix: string,
): string | undefined {
  const declarationName = prefix.length === 0 ? 'xmlns' : `xmlns:${prefix}`;
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

function directChildren(element: XmlElement): XmlElement[] {
  return element.children.filter((child): child is XmlElement => child.type === 'element');
}

function nonNamespaceAttributes(element: XmlElement): XmlAttribute[] {
  return element.attributes.filter(
    ({ name }) => name !== 'xmlns' && !name.startsWith('xmlns:'),
  );
}

function hasNonWhitespaceText(element: XmlElement): boolean {
  return element.children.some(
    (child) => child.type === 'text' && /\S/u.test(child.value),
  );
}
