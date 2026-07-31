import {
  type LosslessXmlDocument,
  type XmlAttribute,
  type XmlElement,
} from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
import {
  SIMPLE_LINE_DASH_CHOICE_NAMES,
  SIMPLE_LINE_FILL_CHOICE_NAMES,
} from './simple-line.internal.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const ARROW_TYPES = new Set<ShapeArrowTypeValue>([
  'none',
  'arrow',
  'diamond',
  'oval',
  'stealth',
  'triangle',
]);
const ARROW_SIZES = new Set(['sm', 'med', 'lg']);
const GEOMETRY_CHOICES = new Set(['prstGeom', 'custGeom']);
const SHAPE_FILL_CHOICES = new Set<string>(SIMPLE_LINE_FILL_CHOICE_NAMES);
const LATER_PROPERTY_CHOICES = new Set([
  'effectLst',
  'effectDag',
  'scene3d',
  'sp3d',
  'extLst',
]);
const LINE_CHILD_STAGES = new Map<string, number>([
  ...SIMPLE_LINE_FILL_CHOICE_NAMES.map((name) => [name, 0] as const),
  ...SIMPLE_LINE_DASH_CHOICE_NAMES.map((name) => [name, 1] as const),
  ['round', 2],
  ['bevel', 2],
  ['miter', 2],
  ['headEnd', 3],
  ['tailEnd', 4],
  ['extLst', 5],
]);

export type ShapeArrowTypeValue =
  | 'none'
  | 'arrow'
  | 'diamond'
  | 'oval'
  | 'stealth'
  | 'triangle';

export interface NormalizedShapeArrows {
  readonly begin?: ShapeArrowTypeValue;
  readonly end?: ShapeArrowTypeValue;
}

interface EndpointState {
  readonly element: XmlElement;
  readonly type: XmlAttribute;
}

interface ExistingArrowState {
  readonly line: XmlElement;
  readonly head: EndpointState | undefined;
  readonly tail: EndpointState | undefined;
  readonly extension: XmlElement | undefined;
  readonly prefix: string;
}

interface LocalEdit {
  readonly start: number;
  readonly end: number;
  readonly replacement: string;
}

export function normalizeShapeArrows(
  value: unknown,
  context: string,
): NormalizedShapeArrows | undefined {
  if (value === undefined) return undefined;
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${context} must be an object`);
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError(`${context} must be an ordinary object`);
  }

  let begin: ShapeArrowTypeValue | undefined;
  let end: ShapeArrowTypeValue | undefined;
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || (key !== 'begin' && key !== 'end')) {
      throw new TypeError(`${context} contains unsupported property ${String(key)}`);
    }
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`${context} property ${key} must be a data property`);
    }
    if (descriptor.value === undefined) continue;
    if (
      typeof descriptor.value !== 'string'
      || !ARROW_TYPES.has(descriptor.value as ShapeArrowTypeValue)
    ) {
      throw new TypeError(`${context} ${key} is unsupported`);
    }
    if (key === 'begin') begin = descriptor.value as ShapeArrowTypeValue;
    else end = descriptor.value as ShapeArrowTypeValue;
  }

  return Object.freeze({
    ...(begin !== undefined ? { begin } : {}),
    ...(end !== undefined ? { end } : {}),
  });
}

export function renderShapeArrows(
  arrows: NormalizedShapeArrows | undefined,
  prefix: string,
): string {
  if (!arrows) return '';
  return (arrows.begin === undefined
    ? ''
    : `<${prefix}headEnd type="${arrows.begin}"/>`)
    + (arrows.end === undefined
      ? ''
      : `<${prefix}tailEnd type="${arrows.end}"/>`);
}

export function readShapeArrows(
  _xml: LosslessXmlDocument,
  shape: XmlElement,
): NormalizedShapeArrows | undefined {
  const properties = resolveShapeProperties(shape);
  if (!properties) return undefined;
  const lines = directChildren(properties).filter(({ localName }) => localName === 'ln');
  if (lines.length !== 1 || namespaceUri(lines[0]!) !== DRAWING_NAMESPACE) {
    return undefined;
  }
  const state = inspectExistingLine(lines[0]!);
  return state ? snapshotFromState(state) : undefined;
}

export function replaceShapeArrows(
  xml: LosslessXmlDocument,
  shape: XmlElement,
  arrows: NormalizedShapeArrows | undefined,
  partUri: string,
): boolean {
  const properties = resolveShapeProperties(shape);
  if (!properties) {
    throw new ModelParseError(
      'Shape must contain exactly one direct shape properties element',
      partUri,
    );
  }

  const lines = directChildren(properties).filter(({ localName }) => localName === 'ln');
  if (lines.length > 1) {
    throw new ModelParseError('Shape contains multiple direct line elements', partUri);
  }
  const line = lines[0];
  if (line && namespaceUri(line) !== DRAWING_NAMESPACE) {
    throw new ModelParseError('Shape line uses an unsafe namespace', partUri);
  }

  if (!line) {
    if (!hasArrow(arrows)) return false;
    const anchor = resolveLineInsertionAnchor(properties, partUri);
    xml.replace(
      anchor.end,
      anchor.end,
      renderLineForParent(arrows!, anchor, properties),
    );
    return true;
  }

  const state = inspectExistingLine(line);
  if (!state) {
    throw new ModelParseError('Shape arrows are not safely editable', partUri);
  }
  const current = snapshotFromState(state);
  if (shapeArrowsEqual(current, arrows)) return false;

  const original = xml.original(line);
  const updated = patchExistingLine(original, state, arrows);
  if (updated === original) return false;
  xml.replaceElement(line, updated);
  return true;
}

export function shapeArrowsEqual(
  left: NormalizedShapeArrows | undefined,
  right: NormalizedShapeArrows | undefined,
): boolean {
  return left?.begin === right?.begin && left?.end === right?.end;
}

function snapshotFromState(
  state: ExistingArrowState,
): NormalizedShapeArrows | undefined {
  if (!state.head && !state.tail) return undefined;
  return Object.freeze({
    ...(state.head
      ? { begin: state.head.type.value as ShapeArrowTypeValue }
      : {}),
    ...(state.tail
      ? { end: state.tail.type.value as ShapeArrowTypeValue }
      : {}),
  });
}

function inspectExistingLine(line: XmlElement): ExistingArrowState | undefined {
  if (namespaceUri(line) !== DRAWING_NAMESPACE) return undefined;
  if (line.children.some((child) => child.type === 'text' && /\S/u.test(child.value))) {
    return undefined;
  }
  const children = directChildren(line);
  let previousStage = -1;
  let head: EndpointState | undefined;
  let tail: EndpointState | undefined;
  let extension: XmlElement | undefined;

  for (const child of children) {
    const stage = LINE_CHILD_STAGES.get(child.localName);
    if (
      stage === undefined
      || namespaceUri(child) !== DRAWING_NAMESPACE
      || stage < previousStage
    ) return undefined;
    previousStage = stage;

    if (
      child.localName !== 'headEnd'
      && child.localName !== 'tailEnd'
      && containsNestedEndpoint(child)
    ) return undefined;

    if (child.localName === 'headEnd') {
      if (head) return undefined;
      head = inspectEndpoint(child);
      if (!head) return undefined;
    } else if (child.localName === 'tailEnd') {
      if (tail) return undefined;
      tail = inspectEndpoint(child);
      if (!tail) return undefined;
    } else if (child.localName === 'extLst') {
      if (extension) return undefined;
      extension = child;
    }
  }

  return {
    line,
    head,
    tail,
    extension,
    prefix: lexicalPrefix(line.name),
  };
}

function inspectEndpoint(element: XmlElement): EndpointState | undefined {
  if (
    namespaceUri(element) !== DRAWING_NAMESPACE
    || directChildren(element).length !== 0
    || element.children.some((child) => child.type === 'text' && /\S/u.test(child.value))
  ) return undefined;

  let type: XmlAttribute | undefined;
  let widthSeen = false;
  let lengthSeen = false;
  for (const attribute of nonNamespaceAttributes(element)) {
    if (attribute.name === 'type') {
      if (type || !ARROW_TYPES.has(attribute.value as ShapeArrowTypeValue)) {
        return undefined;
      }
      type = attribute;
      continue;
    }
    if (attribute.name === 'w') {
      if (widthSeen || !ARROW_SIZES.has(attribute.value)) return undefined;
      widthSeen = true;
      continue;
    }
    if (attribute.name === 'len') {
      if (lengthSeen || !ARROW_SIZES.has(attribute.value)) return undefined;
      lengthSeen = true;
      continue;
    }
    return undefined;
  }
  return type ? { element, type } : undefined;
}

function containsNestedEndpoint(element: XmlElement): boolean {
  for (const child of directChildren(element)) {
    if (
      (child.localName === 'headEnd' || child.localName === 'tailEnd')
      && namespaceUri(child) === DRAWING_NAMESPACE
    ) return true;
    if (containsNestedEndpoint(child)) return true;
  }
  return false;
}

function patchExistingLine(
  source: string,
  state: ExistingArrowState,
  target: NormalizedShapeArrows | undefined,
): string {
  const offset = state.line.start;
  if (state.line.selfClosing) {
    const rendered = renderShapeArrows(target, qualifiedPrefix(state.prefix));
    if (rendered === '') return source;
    const marker = source.lastIndexOf('/>');
    return source.slice(0, marker) + `>${rendered}</${state.line.name}>`;
  }

  const edits: LocalEdit[] = [];
  patchExistingEndpoint(edits, state.head, target?.begin, offset);
  patchExistingEndpoint(edits, state.tail, target?.end, offset);

  const prefix = qualifiedPrefix(state.prefix);
  const newHead = !state.head && target?.begin !== undefined;
  const newTail = !state.tail && target?.end !== undefined;
  if (newHead && newTail) {
    const insertion = baseEndpointInsertion(state);
    edits.push({
      start: insertion - offset,
      end: insertion - offset,
      replacement: renderShapeArrows(target, prefix),
    });
  } else {
    if (newHead) {
      const insertion = state.tail?.element.start ?? baseEndpointInsertion(state);
      edits.push({
        start: insertion - offset,
        end: insertion - offset,
        replacement: `<${prefix}headEnd type="${target!.begin}"/>`,
      });
    }
    if (newTail) {
      const insertion = state.head?.element.end ?? baseEndpointInsertion(state);
      edits.push({
        start: insertion - offset,
        end: insertion - offset,
        replacement: `<${prefix}tailEnd type="${target!.end}"/>`,
      });
    }
  }
  return applyLocalEdits(source, edits);
}

function patchExistingEndpoint(
  edits: LocalEdit[],
  endpoint: EndpointState | undefined,
  target: ShapeArrowTypeValue | undefined,
  offset: number,
): void {
  if (!endpoint) return;
  if (target === undefined) {
    edits.push({
      start: endpoint.element.start - offset,
      end: endpoint.element.end - offset,
      replacement: '',
    });
    return;
  }
  if (endpoint.type.value !== target) {
    edits.push({
      start: endpoint.type.valueStart - offset,
      end: endpoint.type.valueEnd - offset,
      replacement: target,
    });
  }
}

function baseEndpointInsertion(state: ExistingArrowState): number {
  return state.extension?.start ?? state.line.endTagStart;
}

function applyLocalEdits(source: string, edits: readonly LocalEdit[]): string {
  let output = source;
  const ordered = [...edits].sort(
    (left, right) => right.start - left.start || right.end - left.end,
  );
  let previousStart = source.length;
  for (const edit of ordered) {
    if (
      edit.start < 0
      || edit.end < edit.start
      || edit.end > source.length
      || edit.end > previousStart
    ) {
      throw new Error('Overlapping local shape arrow edits');
    }
    output = output.slice(0, edit.start) + edit.replacement + output.slice(edit.end);
    previousStart = edit.start;
  }
  return output;
}

function resolveLineInsertionAnchor(
  properties: XmlElement,
  partUri: string,
): XmlElement {
  const children = directChildren(properties);
  const geometries = children.filter(({ localName }) => GEOMETRY_CHOICES.has(localName));
  if (
    geometries.length !== 1
    || namespaceUri(geometries[0]!) !== DRAWING_NAMESPACE
  ) {
    throw new ModelParseError(
      'Shape must contain exactly one direct geometry before inserting a line',
      partUri,
    );
  }
  const fills = children.filter(({ localName }) => SHAPE_FILL_CHOICES.has(localName));
  if (fills.length > 1 || (fills[0] && namespaceUri(fills[0]) !== DRAWING_NAMESPACE)) {
    throw new ModelParseError('Shape contains unsafe direct fill state', partUri);
  }

  const geometry = geometries[0]!;
  const fill = fills[0];
  const geometryIndex = children.indexOf(geometry);
  const anchor = fill ?? geometry;
  const anchorIndex = children.indexOf(anchor);
  if (fill && anchorIndex <= geometryIndex) {
    throw new ModelParseError('Shape fill is not in a safe line insertion position', partUri);
  }
  const transforms = children.filter(({ localName }) => localName === 'xfrm');
  if (
    transforms.length > 1
    || transforms.some((child) => children.indexOf(child) >= geometryIndex)
  ) {
    throw new ModelParseError('Shape transform is not in a safe line insertion position', partUri);
  }
  if (children.some(
    (child, index) => LATER_PROPERTY_CHOICES.has(child.localName)
      && namespaceUri(child) === DRAWING_NAMESPACE
      && index < anchorIndex,
  )) {
    throw new ModelParseError('Shape properties have an unsafe line insertion order', partUri);
  }
  return anchor;
}

function renderLineForParent(
  arrows: NormalizedShapeArrows,
  prefixSource: XmlElement,
  parent: XmlElement,
): string {
  const prefix = lexicalPrefix(prefixSource.name);
  const qualified = qualifiedPrefix(prefix);
  const namespaceDeclaration = namespaceUriForPrefix(parent, prefix) === DRAWING_NAMESPACE
    ? ''
    : prefix === ''
      ? ` xmlns="${DRAWING_NAMESPACE}"`
      : ` xmlns:${prefix}="${DRAWING_NAMESPACE}"`;
  return `<${qualified}ln${namespaceDeclaration}>${renderShapeArrows(arrows, qualified)}` +
    `</${qualified}ln>`;
}

function resolveShapeProperties(shape: XmlElement): XmlElement | undefined {
  if (shape.localName !== 'sp' || namespaceUri(shape) !== PRESENTATION_NAMESPACE) {
    return undefined;
  }
  const candidates = directChildren(shape).filter(({ localName }) => localName === 'spPr');
  if (candidates.length !== 1) return undefined;
  const properties = candidates[0]!;
  return namespaceUri(properties) === PRESENTATION_NAMESPACE
    ? properties
    : undefined;
}

function hasArrow(arrows: NormalizedShapeArrows | undefined): boolean {
  return arrows?.begin !== undefined || arrows?.end !== undefined;
}

function qualifiedPrefix(prefix: string): string {
  return prefix === '' ? '' : `${prefix}:`;
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

function nonNamespaceAttributes(element: XmlElement): XmlAttribute[] {
  return element.attributes.filter(
    ({ name }) => name !== 'xmlns' && !name.startsWith('xmlns:'),
  );
}

function directChildren(element: XmlElement): XmlElement[] {
  return element.children.filter(
    (child): child is XmlElement => child.type === 'element',
  );
}
