import {
  escapeXmlAttribute,
  type LosslessXmlDocument,
  type XmlElement,
} from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
import {
  PRESET_SHAPE_TYPES,
  type ShapeAdjustment,
} from './preset-shape.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const PRESET_SHAPE_TYPE_SET: ReadonlySet<string> = new Set(PRESET_SHAPE_TYPES);
const ENTRY_KEYS = new Set(['name', 'value']);
const FORMULA_PATTERN = /^val[ \t\r\n]+([+-]?\d+)$/;

export type NormalizedShapeAdjustments = readonly Readonly<ShapeAdjustment>[];

interface AdjustmentOwnerState {
  readonly list: XmlElement;
  readonly prefix: string;
  readonly snapshot: NormalizedShapeAdjustments;
}

export function normalizeShapeAdjustments(
  value: unknown,
  context: string,
): NormalizedShapeAdjustments {
  if (!Array.isArray(value) || Object.getPrototypeOf(value) !== Array.prototype) {
    throw new TypeError(`${context} must be an ordinary array`);
  }
  const keys = Reflect.ownKeys(value);
  if (keys.length !== value.length + 1 || !keys.includes('length')) {
    throw new TypeError(`${context} must be a dense array without extra properties`);
  }

  const names = new Set<string>();
  const result: ShapeAdjustment[] = [];
  for (let index = 0; index < value.length; index += 1) {
    const key = String(index);
    if (!keys.includes(key)) {
      throw new TypeError(`${context} must be a dense array`);
    }
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`${context} item ${index} must be a data property`);
    }
    const candidate = readEntry(descriptor.value, `${context} item ${index}`);
    const name = candidate.name;
    if (typeof name !== 'string' || name.length === 0) {
      throw new TypeError(`${context} item ${index} name must be a non-empty string`);
    }
    if (containsInvalidXmlCharacter(name)) {
      throw new TypeError(`${context} item ${index} name contains invalid XML characters`);
    }
    if (names.has(name)) {
      throw new TypeError(`${context} contains duplicate adjustment name ${name}`);
    }
    const adjustmentValue = candidate.value;
    if (typeof adjustmentValue !== 'number' || !Number.isFinite(adjustmentValue)) {
      throw new TypeError(`${context} item ${index} value must be finite`);
    }
    if (!Number.isSafeInteger(adjustmentValue)) {
      throw new RangeError(`${context} item ${index} value must be a safe integer`);
    }
    names.add(name);
    result.push(Object.freeze({ name, value: normalizeNegativeZero(adjustmentValue) }));
  }
  return Object.freeze(result);
}

export function renderShapeAdjustmentList(
  adjustments: NormalizedShapeAdjustments,
  prefix: string,
): string {
  if (adjustments.length === 0) return `<${prefix}avLst/>`;
  const guides = adjustments.map(({ name, value }) =>
    `<${prefix}gd name="${escapeXmlAttribute(name)}" fmla="val ${value}"/>`,
  ).join('');
  return `<${prefix}avLst>${guides}</${prefix}avLst>`;
}

export function readShapeAdjustments(
  _xml: LosslessXmlDocument,
  shape: XmlElement,
): NormalizedShapeAdjustments | undefined {
  return inspectAdjustmentOwner(shape)?.snapshot;
}

export function replaceShapeAdjustments(
  xml: LosslessXmlDocument,
  shape: XmlElement,
  adjustments: NormalizedShapeAdjustments,
  partUri: string,
): boolean {
  const state = inspectAdjustmentOwner(shape);
  if (!state) {
    throw new ModelParseError('Shape adjustment state is not safely editable', partUri);
  }
  if (shapeAdjustmentsEqual(state.snapshot, adjustments)) return false;

  const qualified = qualifiedPrefix(state.prefix);
  let replacement = renderShapeAdjustmentList(adjustments, qualified);
  const parentBinding = state.list.parent
    ? namespaceUriForPrefix(state.list.parent, state.prefix)
    : undefined;
  if (parentBinding !== DRAWING_NAMESPACE) {
    const declaration = state.prefix === ''
      ? ` xmlns="${DRAWING_NAMESPACE}"`
      : ` xmlns:${state.prefix}="${DRAWING_NAMESPACE}"`;
    replacement = replacement.replace(
      `<${qualified}avLst`,
      `<${qualified}avLst${declaration}`,
    );
  }
  xml.replaceElement(state.list, replacement);
  return true;
}

export function shapeAdjustmentsEqual(
  left: NormalizedShapeAdjustments | undefined,
  right: NormalizedShapeAdjustments | undefined,
): boolean {
  if (left === undefined || right === undefined) return left === right;
  if (left.length !== right.length) return false;
  return left.every((entry, index) => {
    const other = right[index];
    return other !== undefined
      && entry.name === other.name
      && entry.value === other.value;
  });
}

function inspectAdjustmentOwner(shape: XmlElement): AdjustmentOwnerState | undefined {
  if (shape.localName !== 'sp' || namespaceUri(shape) !== PRESENTATION_NAMESPACE) {
    return undefined;
  }
  const properties = directChildren(shape).filter(({ localName }) => localName === 'spPr');
  if (properties.length !== 1 || namespaceUri(properties[0]!) !== PRESENTATION_NAMESPACE) {
    return undefined;
  }
  const geometries = directChildren(properties[0]!)
    .filter(({ localName }) => localName === 'prstGeom' || localName === 'custGeom');
  const geometry = geometries[0];
  if (
    geometries.length !== 1
    || !geometry
    || geometry.localName !== 'prstGeom'
    || namespaceUri(geometry) !== DRAWING_NAMESPACE
    || !hasCanonicalPresetType(geometry)
  ) return undefined;

  const lists = directChildren(geometry).filter(({ localName }) => localName === 'avLst');
  const list = lists[0];
  if (
    lists.length !== 1
    || !list
    || namespaceUri(list) !== DRAWING_NAMESPACE
    || nonNamespaceAttributes(list).length !== 0
    || hasNonWhitespaceText(list)
  ) return undefined;

  const names = new Set<string>();
  const snapshot: ShapeAdjustment[] = [];
  for (const guide of directChildren(list)) {
    if (
      guide.localName !== 'gd'
      || namespaceUri(guide) !== DRAWING_NAMESPACE
      || directChildren(guide).length !== 0
      || hasNonWhitespaceText(guide)
    ) return undefined;
    const attributes = nonNamespaceAttributes(guide);
    if (attributes.length !== 2) return undefined;
    const nameAttribute = attributes.find(({ name }) => name === 'name');
    const formulaAttribute = attributes.find(({ name }) => name === 'fmla');
    if (!nameAttribute || !formulaAttribute) return undefined;
    const name = nameAttribute.value;
    if (
      name.length === 0
      || containsInvalidXmlCharacter(name)
      || names.has(name)
    ) return undefined;
    const match = FORMULA_PATTERN.exec(formulaAttribute.value);
    if (!match?.[1]) return undefined;
    const value = Number(match[1]);
    if (!Number.isSafeInteger(value)) return undefined;
    names.add(name);
    snapshot.push(Object.freeze({ name, value: normalizeNegativeZero(value) }));
  }

  return {
    list,
    prefix: lexicalPrefix(list.name),
    snapshot: Object.freeze(snapshot),
  };
}

function hasCanonicalPresetType(geometry: XmlElement): boolean {
  const attributes = geometry.attributes.filter(
    ({ localName, name }) => localName === 'prst' && !name.startsWith('xmlns:'),
  );
  return attributes.length === 1
    && attributes[0]?.name === 'prst'
    && PRESET_SHAPE_TYPE_SET.has(attributes[0].value);
}

function readEntry(value: unknown, context: string): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${context} must be an object`);
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError(`${context} must be an ordinary object`);
  }
  const result: Record<string, unknown> = {};
  const keys = Reflect.ownKeys(value);
  if (keys.length !== ENTRY_KEYS.size) {
    throw new TypeError(`${context} must contain only name and value`);
  }
  for (const key of keys) {
    if (typeof key !== 'string' || !ENTRY_KEYS.has(key)) {
      throw new TypeError(`${context} contains unsupported property ${String(key)}`);
    }
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`${context} property ${key} must be a data property`);
    }
    result[key] = descriptor.value;
  }
  if (!Object.hasOwn(result, 'name') || !Object.hasOwn(result, 'value')) {
    throw new TypeError(`${context} must contain name and value`);
  }
  return result;
}

function normalizeNegativeZero(value: number): number {
  return Object.is(value, -0) ? 0 : value;
}

function containsInvalidXmlCharacter(value: string): boolean {
  for (const character of value) {
    const codePoint = character.codePointAt(0)!;
    if (
      codePoint !== 0x9
      && codePoint !== 0xA
      && codePoint !== 0xD
      && (codePoint < 0x20 || codePoint > 0xD7FF)
      && (codePoint < 0xE000 || codePoint > 0xFFFD)
      && (codePoint < 0x10000 || codePoint > 0x10FFFF)
    ) return true;
  }
  return false;
}

function nonNamespaceAttributes(element: XmlElement) {
  return element.attributes.filter(
    ({ name }) => name !== 'xmlns' && !name.startsWith('xmlns:'),
  );
}

function hasNonWhitespaceText(element: XmlElement): boolean {
  return element.children.some(
    (child) => child.type === 'text' && /[^ \t\r\n]/.test(child.value),
  );
}

function directChildren(element: XmlElement): XmlElement[] {
  return element.children.filter(
    (child): child is XmlElement => child.type === 'element',
  );
}

function namespaceUri(element: XmlElement): string | undefined {
  return namespaceUriForPrefix(element, lexicalPrefix(element.name));
}

function namespaceUriForPrefix(
  element: XmlElement,
  prefix: string,
): string | undefined {
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

function qualifiedPrefix(prefix: string): string {
  return prefix === '' ? '' : `${prefix}:`;
}
