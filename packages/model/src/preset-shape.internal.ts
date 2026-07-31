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

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const EMU_PER_INCH = 914_400;
const MAX_ROTATION = 21_600_000;
const PRESET_SHAPE_TYPE_SET: ReadonlySet<string> = new Set(PRESET_SHAPE_TYPES);
const OPTION_KEYS = new Set([
  'name',
  'adjustments',
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

export interface NormalizedPresetShape {
  readonly type: PresetShapeType;
  readonly name: string | undefined;
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

interface ResolvedPresetGeometry {
  readonly geometry: XmlElement;
  readonly type: PresetShapeType;
  readonly prefix: string;
}

export function normalizePresetShape(
  type: unknown,
  options: unknown = undefined,
): NormalizedPresetShape {
  if (typeof type !== 'string' || !PRESET_SHAPE_TYPE_SET.has(type)) {
    throw new TypeError('Preset shape type must be a canonical preset shape string');
  }
  const values = readOptions(options);
  const name = values.name;
  if (name !== undefined) {
    if (typeof name !== 'string') {
      throw new TypeError('Preset shape name must be a string');
    }
    if (containsInvalidXmlCharacter(name)) {
      throw new TypeError('Preset shape name contains invalid XML characters');
    }
  }
  const width = normalizeNumber(values.width, EMU_PER_INCH, 'width');
  const height = normalizeNumber(values.height, EMU_PER_INCH, 'height');
  if (width <= 0) throw new RangeError('Preset shape width must be greater than zero');
  if (height <= 0) throw new RangeError('Preset shape height must be greater than zero');
  const rotation = normalizeNumber(values.rotation, 0, 'rotation');
  if (rotation < -MAX_ROTATION || rotation > MAX_ROTATION) {
    throw new RangeError('Preset shape rotation must be between -21600000 and 21600000');
  }
  const fill = normalizeSimpleFill(values.fill, 'Preset shape fill') ?? { kind: 'none' };
  const line = normalizeSimpleLine(values.line, 'Preset shape line');
  const arrows = normalizeShapeArrows(values.arrows, 'Preset shape arrows');
  const hyperlink = values.hyperlink === undefined
    ? undefined
    : normalizeHyperlink(values.hyperlink, 'Preset shape hyperlink');
  const shadow = values.shadow === undefined
    ? undefined
    : normalizeShapeShadow(values.shadow, 'Preset shape shadow');
  const adjustments = normalizeShapeAdjustments(
    values.adjustments === undefined ? [] : values.adjustments,
    'Preset shape adjustments',
  );

  return Object.freeze({
    type: type as PresetShapeType,
    name: name as string | undefined,
    adjustments,
    fill,
    line,
    arrows,
    hyperlink,
    shadow,
    x: normalizeNumber(values.x, EMU_PER_INCH, 'x'),
    y: normalizeNumber(values.y, EMU_PER_INCH, 'y'),
    width,
    height,
    rotation,
    flipHorizontal: normalizeBoolean(values.flipHorizontal, false, 'flipHorizontal'),
    flipVertical: normalizeBoolean(values.flipVertical, false, 'flipVertical'),
  });
}

export function renderPresetShapeXml(
  id: number,
  shape: NormalizedPresetShape,
  hyperlinkRelationshipId?: string,
): string {
  if ((shape.hyperlink === undefined) !== (hyperlinkRelationshipId === undefined)) {
    throw new TypeError('Preset shape hyperlink and relationship ID must be supplied together');
  }
  const name = escapeXmlAttribute(shape.name ?? `Shape ${id}`);
  const type = escapeXmlAttribute(shape.type);
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
  return `<p:sp xmlns:p="${PRESENTATION_NAMESPACE}" xmlns:a="${DRAWING_NAMESPACE}">` +
    `<p:nvSpPr>${nonVisualProperties}<p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm${transformAttributes}><a:off x="${shape.x}" y="${shape.y}"/>` +
    `<a:ext cx="${shape.width}" cy="${shape.height}"/></a:xfrm>` +
    `<a:prstGeom prst="${type}">` +
    `${renderShapeAdjustmentList(shape.adjustments, 'a:')}</a:prstGeom>` +
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
  return `<a:ln w="${Math.round(line.width * 12_700)}">` +
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
  const state = resolvePresetGeometry(shape);
  if (!state) {
    throw new ModelParseError('Shape preset geometry is not safely editable', partUri);
  }
  if (state.type === type) return false;
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
  if (
    shape.localName !== 'sp'
    || namespaceUri(shape) !== PRESENTATION_NAMESPACE
  ) return undefined;
  const properties = directChildren(shape).filter(({ localName }) => localName === 'spPr');
  if (
    properties.length !== 1
    || namespaceUri(properties[0]!) !== PRESENTATION_NAMESPACE
  ) return undefined;
  const geometries = directChildren(properties[0]!)
    .filter(({ localName }) => localName === 'prstGeom');
  const geometry = geometries[0];
  if (
    geometries.length !== 1
    || !geometry
    || namespaceUri(geometry) !== DRAWING_NAMESPACE
  ) return undefined;
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

function readOptions(value: unknown): Record<string, unknown> {
  if (value === undefined) return {};
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError('Preset shape options must be an object');
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError('Preset shape options must be an ordinary object');
  }
  const result: Record<string, unknown> = {};
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !OPTION_KEYS.has(key)) {
      throw new TypeError(`Preset shape options contain unsupported property ${String(key)}`);
    }
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`Preset shape option ${key} must be a data property`);
    }
    result[key] = descriptor.value;
  }
  return result;
}

function normalizeNumber(
  value: unknown,
  defaultValue: number,
  name: string,
): number {
  if (value === undefined) return defaultValue;
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`Preset shape ${name} must be finite`);
  }
  const rounded = Math.round(value);
  if (!Number.isSafeInteger(rounded)) {
    throw new RangeError(`Preset shape ${name} must round to a safe integer`);
  }
  return rounded;
}

function normalizeBoolean(
  value: unknown,
  defaultValue: boolean,
  name: string,
): boolean {
  if (value === undefined) return defaultValue;
  if (typeof value !== 'boolean') {
    throw new TypeError(`Preset shape ${name} must be a boolean`);
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
