import {
  escapeXmlAttribute,
  type LosslessXmlDocument,
  type XmlElement,
} from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
import {
  PRESET_SHAPE_TYPES,
  type PresetShapeType,
  type ShapeFill,
} from './preset-shape.js';
import {
  normalizeShapeAdjustments,
  renderShapeAdjustmentList,
  type NormalizedShapeAdjustments,
} from './shape-adjustments.internal.js';
import {
  normalizeSimpleFill,
  renderSimpleFill,
} from './simple-fill.internal.js';
import {
  normalizeSimpleLine,
  renderSimpleLine,
  renderSimpleLineAttributes,
  type NormalizedSimpleLine,
} from './simple-line.internal.js';
import {
  normalizeShapeArrows,
  renderShapeArrows,
  type NormalizedShapeArrows,
} from './shape-arrows.internal.js';
import {
  normalizeHyperlink,
  renderShapeHyperlink,
  type NormalizedHyperlink,
} from './shape-hyperlink.internal.js';
import {
  normalizeShapeShadow,
  renderSimpleShadow,
  type NormalizedShapeShadow,
} from './simple-shadow.internal.js';
import {
  normalizeCustomGeometry,
  readCustomGeometry,
  renderCustomGeometry,
  type NormalizedCustomGeometry,
} from './custom-geometry.internal.js';
import {
  normalizePlaceholderSelector,
} from './placeholder.internal.js';
import type { PlaceholderIdentity, PlaceholderSelector } from './placeholder.js';
import { resolveSlideCoordinate } from './slide-coordinate.internal.js';
import type { Emu, SlideSize } from './units.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const EMU_PER_INCH = 914_400;
const MAX_ROTATION = 21_600_000;
const PRESET_SHAPE_TYPE_SET: ReadonlySet<string> = new Set(PRESET_SHAPE_TYPES);
const EMPTY_SHAPE_ADJUSTMENTS: NormalizedShapeAdjustments = Object.freeze([]);
const COMMON_OPTION_KEYS = new Set([
  'name',
  'placeholder',
  'fill',
  'line',
  'arrows',
  'hyperlink',
  'shadow',
  'x',
  'y',
  'width',
  'height',
  'rotation',
  'flipHorizontal',
  'flipVertical',
]);
const PRESET_OPTION_KEYS = new Set([...COMMON_OPTION_KEYS, 'adjustments']);

export interface NormalizedPresetShape {
  readonly type: PresetShapeType;
  readonly name: string | undefined;
  readonly placeholder?: PlaceholderSelector;
  readonly adjustments: NormalizedShapeAdjustments;
  readonly fill: ShapeFill;
  readonly line: NormalizedSimpleLine | undefined;
  readonly arrows: NormalizedShapeArrows | undefined;
  readonly hyperlink: NormalizedHyperlink | undefined;
  readonly shadow: NormalizedShapeShadow | undefined;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly rotation: number;
  readonly flipHorizontal: boolean;
  readonly flipVertical: boolean;
}

export interface NormalizedCustomShape {
  readonly geometry: NormalizedCustomGeometry;
  readonly name: string | undefined;
  readonly placeholder?: PlaceholderSelector;
  readonly fill: ShapeFill;
  readonly line: NormalizedSimpleLine | undefined;
  readonly arrows: NormalizedShapeArrows | undefined;
  readonly hyperlink: NormalizedHyperlink | undefined;
  readonly shadow: NormalizedShapeShadow | undefined;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly rotation: number;
  readonly flipHorizontal: boolean;
  readonly flipVertical: boolean;
}

interface NormalizedShapeOptions {
  readonly name: string | undefined;
  readonly placeholder?: PlaceholderSelector;
  readonly fill: ShapeFill;
  readonly line: NormalizedSimpleLine | undefined;
  readonly arrows: NormalizedShapeArrows | undefined;
  readonly hyperlink: NormalizedHyperlink | undefined;
  readonly shadow: NormalizedShapeShadow | undefined;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly rotation: number;
  readonly flipHorizontal: boolean;
  readonly flipVertical: boolean;
}

interface ResolvedPresetGeometry {
  readonly geometry: XmlElement;
  readonly type: PresetShapeType;
  readonly prefix: string;
}

export function normalizePresetShape(
  type: unknown,
  options: unknown = undefined,
  slideSize?: Readonly<SlideSize>,
): NormalizedPresetShape {
  const normalizedType = normalizePresetShapeType(type, 'Preset shape type');
  const values = readOptions(options, PRESET_OPTION_KEYS, 'Preset shape');
  const normalizedOptions = normalizeShapeOptions(values, 'Preset shape', slideSize);
  const adjustments = normalizeShapeAdjustments(
    values.adjustments === undefined ? [] : values.adjustments,
    'Preset shape adjustments',
  );

  return Object.freeze({
    type: normalizedType,
    ...normalizedOptions,
    adjustments,
  });
}

export function normalizePresetShapeType(
  value: unknown,
  context: string,
): PresetShapeType {
  if (typeof value !== 'string' || !PRESET_SHAPE_TYPE_SET.has(value)) {
    throw new TypeError(`${context} must be a canonical preset shape string`);
  }
  return value as PresetShapeType;
}

export function normalizeCustomShape(
  geometry: unknown,
  options: unknown = undefined,
  slideSize?: Readonly<SlideSize>,
): NormalizedCustomShape {
  const normalizedGeometry = normalizeCustomGeometry(geometry, 'Custom geometry');
  const values = readOptions(options, COMMON_OPTION_KEYS, 'Custom shape');
  return Object.freeze({
    geometry: normalizedGeometry,
    ...normalizeShapeOptions(values, 'Custom shape', slideSize),
  });
}

function normalizeShapeOptions(
  values: Record<string, unknown>,
  context: string,
  slideSize: Readonly<SlideSize> | undefined,
): NormalizedShapeOptions {
  const name = values.name;
  if (name !== undefined) {
    if (typeof name !== 'string') {
      throw new TypeError(`${context} name must be a string`);
    }
    if (containsInvalidXmlCharacter(name)) {
      throw new TypeError(`${context} name contains invalid XML characters`);
    }
  }
  const width = resolveSlideCoordinate(
    values.width,
    'horizontal',
    slideSize,
    EMU_PER_INCH as Emu,
    `${context} width`,
  );
  const height = resolveSlideCoordinate(
    values.height,
    'vertical',
    slideSize,
    EMU_PER_INCH as Emu,
    `${context} height`,
  );
  if (width <= 0) throw new RangeError(`${context} width must be greater than zero`);
  if (height <= 0) throw new RangeError(`${context} height must be greater than zero`);
  const rotation = normalizeNumber(values.rotation, 0, 'rotation', context);
  if (rotation < -MAX_ROTATION || rotation > MAX_ROTATION) {
    throw new RangeError(`${context} rotation must be between -21600000 and 21600000`);
  }
  const fill = normalizeSimpleFill(values.fill, `${context} fill`) ?? { kind: 'none' };
  const line = normalizeSimpleLine(values.line, `${context} line`);
  const arrows = normalizeShapeArrows(values.arrows, `${context} arrows`);
  const hyperlink = values.hyperlink === undefined
    ? undefined
    : normalizeHyperlink(values.hyperlink, `${context} hyperlink`);
  const shadow = values.shadow === undefined
    ? undefined
    : normalizeShapeShadow(values.shadow, `${context} shadow`);
  const placeholder = values.placeholder === undefined
    ? undefined
    : normalizePlaceholderSelector(values.placeholder);

  return {
    name: name as string | undefined,
    ...(placeholder === undefined ? {} : { placeholder }),
    fill,
    line,
    arrows,
    hyperlink,
    shadow,
    x: resolveSlideCoordinate(
      values.x,
      'horizontal',
      slideSize,
      EMU_PER_INCH as Emu,
      `${context} x`,
    ),
    y: resolveSlideCoordinate(
      values.y,
      'vertical',
      slideSize,
      EMU_PER_INCH as Emu,
      `${context} y`,
    ),
    width,
    height,
    rotation,
    flipHorizontal: normalizeBoolean(values.flipHorizontal, false, 'flipHorizontal', context),
    flipVertical: normalizeBoolean(values.flipVertical, false, 'flipVertical', context),
  };
}

export function renderPresetShapeXml(
  id: number,
  shape: NormalizedPresetShape,
  hyperlinkRelationshipId?: string,
  placeholder?: Readonly<PlaceholderIdentity>,
): string {
  return renderShapeXml(
    id,
    shape,
    renderPresetShapeGeometry(shape.type, 'a:', shape.adjustments),
    hyperlinkRelationshipId,
    'Preset shape',
    placeholder,
  );
}

export function renderPresetShapeGeometry(
  type: PresetShapeType,
  prefix = 'a:',
  adjustments: NormalizedShapeAdjustments = EMPTY_SHAPE_ADJUSTMENTS,
): string {
  return `<${prefix}prstGeom prst="${escapeXmlAttribute(type)}">` +
    `${renderShapeAdjustmentList(adjustments, prefix)}</${prefix}prstGeom>`;
}

export function renderCustomShapeXml(
  id: number,
  shape: NormalizedCustomShape,
  hyperlinkRelationshipId?: string,
  placeholder?: Readonly<PlaceholderIdentity>,
): string {
  return renderShapeXml(
    id,
    shape,
    renderCustomGeometry(shape.geometry, 'a:'),
    hyperlinkRelationshipId,
    'Custom shape',
    placeholder,
  );
}

function renderShapeXml(
  id: number,
  shape: NormalizedShapeOptions,
  geometry: string,
  hyperlinkRelationshipId: string | undefined,
  context: string,
  placeholder?: Readonly<PlaceholderIdentity>,
): string {
  if ((shape.hyperlink === undefined) !== (hyperlinkRelationshipId === undefined)) {
    throw new TypeError(`${context} hyperlink and relationship ID must be supplied together`);
  }
  const name = escapeXmlAttribute(shape.name ?? `Shape ${id}`);
  const hyperlink = shape.hyperlink === undefined
    ? ''
    : renderPresetShapeHyperlink(shape.hyperlink, hyperlinkRelationshipId!);
  const nonVisualProperties = hyperlink === ''
    ? `<p:cNvPr id="${id}" name="${name}"/>`
    : `<p:cNvPr id="${id}" name="${name}">${hyperlink}</p:cNvPr>`;
  const transformAttributes = [
    shape.rotation === 0 ? '' : ` rot="${shape.rotation}"`,
    shape.flipHorizontal ? ' flipH="1"' : '',
    shape.flipVertical ? ' flipV="1"' : '',
  ].join('');
  const effect = shape.shadow === undefined
    ? ''
    : `<a:effectLst>${renderSimpleShadow(shape.shadow, 'a:')}</a:effectLst>`;
  const applicationProperties = placeholder === undefined
    ? '<p:nvPr/>'
    : `<p:nvPr><p:ph type="${placeholder.type}" idx="${placeholder.index}"/></p:nvPr>`;
  return `<p:sp xmlns:p="${PRESENTATION_NAMESPACE}" xmlns:a="${DRAWING_NAMESPACE}">` +
    `<p:nvSpPr>${nonVisualProperties}<p:cNvSpPr/>${applicationProperties}</p:nvSpPr>` +
    `<p:spPr><a:xfrm${transformAttributes}><a:off x="${shape.x}" y="${shape.y}"/>` +
    `<a:ext cx="${shape.width}" cy="${shape.height}"/></a:xfrm>` +
    geometry +
    `${renderSimpleFill(shape.fill, 'a:')}${renderPresetLine(shape.line, shape.arrows)}` +
    effect +
    '</p:spPr></p:sp>';
}

function renderPresetShapeHyperlink(
  hyperlink: NormalizedHyperlink,
  relationshipId: string,
): string {
  const rendered = renderShapeHyperlink(
    hyperlink,
    relationshipId,
    { drawing: 'a', relationship: 'r' },
  );
  const insertionPoint = rendered.indexOf(' ');
  return rendered.slice(0, insertionPoint) +
    ` xmlns:r="${RELATIONSHIP_NAMESPACE}"` +
    rendered.slice(insertionPoint);
}

function renderPresetLine(
  line: NormalizedSimpleLine | undefined,
  arrows: NormalizedShapeArrows | undefined,
): string {
  const arrowXml = renderShapeArrows(arrows, 'a:');
  if (line === undefined) {
    return arrowXml === '' ? '<a:ln/>' : `<a:ln>${arrowXml}</a:ln>`;
  }
  if (line.kind === 'none') {
    return `<a:ln>${renderSimpleLine(line, 'a:')}${arrowXml}</a:ln>`;
  }
  return `<a:ln w="${Math.round(line.width * 12_700)}"${renderSimpleLineAttributes(line)}>` +
    `${renderSimpleLine(line, 'a:')}${arrowXml}</a:ln>`;
}

export function readPresetShapeType(
  _xml: LosslessXmlDocument,
  shape: XmlElement,
): PresetShapeType | undefined {
  return resolvePresetGeometry(shape)?.type;
}

export function replacePresetShapeType(
  xml: LosslessXmlDocument,
  shape: XmlElement,
  type: PresetShapeType,
  partUri: string,
): boolean {
  const presetState = resolvePresetGeometry(shape);
  const state = presetState ?? resolveSupportedCustomGeometry(xml, shape);
  if (!state) {
    throw new ModelParseError('Shape preset geometry is not safely editable', partUri);
  }
  if (presetState?.type === type) return false;
  const prefix = state.prefix === '' ? '' : `${state.prefix}:`;
  const parentBinding = state.geometry.parent
    ? namespaceUriForPrefix(state.geometry.parent, state.prefix)
    : undefined;
  const namespaceDeclaration = parentBinding === DRAWING_NAMESPACE
    ? ''
    : state.prefix === ''
      ? ` xmlns="${DRAWING_NAMESPACE}"`
      : ` xmlns:${state.prefix}="${DRAWING_NAMESPACE}"`;
  xml.replaceElement(
    state.geometry,
    `<${prefix}prstGeom${namespaceDeclaration} prst="${escapeXmlAttribute(type)}">` +
      `<${prefix}avLst/></${prefix}prstGeom>`,
  );
  return true;
}

function resolvePresetGeometry(shape: XmlElement): ResolvedPresetGeometry | undefined {
  const geometry = resolveGeometryChoice(shape);
  if (!geometry || geometry.localName !== 'prstGeom') return undefined;
  const attributes = geometry.attributes.filter(
    ({ name, localName }) => localName === 'prst' && !name.startsWith('xmlns:'),
  );
  const attribute = attributes[0];
  if (
    attributes.length !== 1
    || !attribute
    || attribute.name !== 'prst'
    || !PRESET_SHAPE_TYPE_SET.has(attribute.value)
  ) return undefined;
  return {
    geometry,
    type: attribute.value as PresetShapeType,
    prefix: lexicalPrefix(geometry.name),
  };
}

function resolveSupportedCustomGeometry(
  xml: LosslessXmlDocument,
  shape: XmlElement,
): Omit<ResolvedPresetGeometry, 'type'> | undefined {
  const geometry = resolveGeometryChoice(shape);
  if (
    !geometry
    || geometry.localName !== 'custGeom'
    || readCustomGeometry(xml, shape) === undefined
  ) return undefined;
  return { geometry, prefix: lexicalPrefix(geometry.name) };
}

function resolveGeometryChoice(shape: XmlElement): XmlElement | undefined {
  if (
    shape.localName !== 'sp'
    || namespaceUri(shape) !== PRESENTATION_NAMESPACE
  ) return undefined;
  const properties = directChildren(shape).filter(({ localName }) => localName === 'spPr');
  if (
    properties.length !== 1
    || namespaceUri(properties[0]!) !== PRESENTATION_NAMESPACE
  ) return undefined;
  const geometries = directChildren(properties[0]!).filter(
    ({ localName }) => localName === 'prstGeom' || localName === 'custGeom',
  );
  const geometry = geometries[0];
  if (
    geometries.length !== 1
    || !geometry
    || namespaceUri(geometry) !== DRAWING_NAMESPACE
  ) return undefined;
  return geometry;
}

function readOptions(
  value: unknown,
  allowedKeys: ReadonlySet<string>,
  context: string,
): Record<string, unknown> {
  if (value === undefined) return {};
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${context} options must be an object`);
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError(`${context} options must be an ordinary object`);
  }
  const result: Record<string, unknown> = {};
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !allowedKeys.has(key)) {
      throw new TypeError(`${context} options contain unsupported property ${String(key)}`);
    }
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`${context} option ${key} must be a data property`);
    }
    result[key] = descriptor.value;
  }
  return result;
}

function normalizeNumber(
  value: unknown,
  defaultValue: number,
  name: string,
  context: string,
): number {
  if (value === undefined) return defaultValue;
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} ${name} must be finite`);
  }
  const rounded = Math.round(value);
  if (!Number.isSafeInteger(rounded)) {
    throw new RangeError(`${context} ${name} must round to a safe integer`);
  }
  return rounded;
}

function normalizeBoolean(
  value: unknown,
  defaultValue: boolean,
  name: string,
  context: string,
): boolean {
  if (value === undefined) return defaultValue;
  if (typeof value !== 'boolean') {
    throw new TypeError(`${context} ${name} must be a boolean`);
  }
  return value;
}

function containsInvalidXmlCharacter(value: string): boolean {
  return /[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(value);
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

function directChildren(element: XmlElement): XmlElement[] {
  return element.children.filter(
    (child): child is XmlElement => child.type === 'element',
  );
}
