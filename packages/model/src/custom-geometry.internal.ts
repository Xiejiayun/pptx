import type { LosslessXmlDocument, XmlElement } from '@pptx/lossless-xml';
import type {
  CustomGeometry,
  CustomGeometryCommand,
  CustomGeometryPath,
  CustomGeometryPathFill,
  CustomGeometryPoint,
} from './custom-geometry.js';
import { ModelParseError } from './errors.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const ROOT_KEYS = new Set(['paths']);
const PATH_KEYS = new Set([
  'width',
  'height',
  'fill',
  'stroke',
  'extrusionOk',
  'commands',
]);
const POINT_KEYS = new Set(['x', 'y']);
const COMMAND_KEYS = new Map<string, ReadonlySet<string>>([
  ['moveTo', new Set(['kind', 'point'])],
  ['lineTo', new Set(['kind', 'point'])],
  [
    'arcTo',
    new Set([
      'kind',
      'widthRadius',
      'heightRadius',
      'startAngle',
      'sweepAngle',
    ]),
  ],
  ['quadraticBezierTo', new Set(['kind', 'control', 'end'])],
  ['cubicBezierTo', new Set(['kind', 'control1', 'control2', 'end'])],
  ['close', new Set(['kind'])],
]);
const PATH_FILLS = new Set<CustomGeometryPathFill>([
  'none',
  'norm',
  'lighten',
  'lightenLess',
  'darken',
  'darkenLess',
]);
const CUSTOM_CHILD_STAGES = new Map([
  ['avLst', 0],
  ['gdLst', 1],
  ['ahLst', 2],
  ['cxnLst', 3],
  ['rect', 4],
  ['pathLst', 5],
]);
const EMPTY_LIST_NAMES = ['avLst', 'gdLst', 'ahLst', 'cxnLst'] as const;
const INTEGER_PATTERN = /^[+-]?\d+$/;

export type NormalizedCustomGeometry = Readonly<CustomGeometry>;

interface CustomGeometryOwnerState {
  readonly geometry: XmlElement;
  readonly prefix: string;
  readonly snapshot: NormalizedCustomGeometry;
}

export function normalizeCustomGeometry(
  value: unknown,
  context: string,
): NormalizedCustomGeometry {
  const root = readObject(value, ROOT_KEYS, ROOT_KEYS, context);
  const paths = readArray(root.paths, `${context} paths`);
  if (paths.length === 0) throw new RangeError(`${context} paths must not be empty`);
  const normalizedPaths = paths.map((path, index) =>
    normalizePath(path, `${context} path ${index}`));
  return Object.freeze({ paths: Object.freeze(normalizedPaths) });
}

export function renderCustomGeometry(
  geometry: NormalizedCustomGeometry,
  prefix: string,
): string {
  const paths = geometry.paths.map((path) => renderPath(path, prefix)).join('');
  return `<${prefix}custGeom><${prefix}avLst/><${prefix}gdLst/>` +
    `<${prefix}ahLst/><${prefix}cxnLst/>` +
    `<${prefix}rect l="l" t="t" r="r" b="b"/>` +
    `<${prefix}pathLst>${paths}</${prefix}pathLst></${prefix}custGeom>`;
}

export function readCustomGeometry(
  _xml: LosslessXmlDocument,
  shape: XmlElement,
): NormalizedCustomGeometry | undefined {
  return inspectCustomGeometryOwner(shape)?.snapshot;
}

export function replaceCustomGeometry(
  xml: LosslessXmlDocument,
  shape: XmlElement,
  geometry: NormalizedCustomGeometry,
  partUri: string,
): boolean {
  const state = inspectCustomGeometryOwner(shape);
  if (!state) {
    throw new ModelParseError('Shape custom geometry is not safely editable', partUri);
  }
  if (customGeometryEqual(state.snapshot, geometry)) return false;

  const qualified = qualifiedPrefix(state.prefix);
  let replacement = renderCustomGeometry(geometry, qualified);
  const parentBinding = state.geometry.parent
    ? namespaceUriForPrefix(state.geometry.parent, state.prefix)
    : undefined;
  if (parentBinding !== DRAWING_NAMESPACE) {
    const declaration = state.prefix === ''
      ? ` xmlns="${DRAWING_NAMESPACE}"`
      : ` xmlns:${state.prefix}="${DRAWING_NAMESPACE}"`;
    replacement = replacement.replace(
      `<${qualified}custGeom`,
      `<${qualified}custGeom${declaration}`,
    );
  }
  xml.replaceElement(state.geometry, replacement);
  return true;
}

export function customGeometryEqual(
  left: NormalizedCustomGeometry | undefined,
  right: NormalizedCustomGeometry | undefined,
): boolean {
  if (left === undefined || right === undefined) return left === right;
  if (left.paths.length !== right.paths.length) return false;
  return left.paths.every((path, index) => {
    const other = right.paths[index];
    if (!other) return false;
    if (
      path.width !== other.width
      || path.height !== other.height
      || !optionalPropertyEqual(path, other, 'fill')
      || !optionalPropertyEqual(path, other, 'stroke')
      || !optionalPropertyEqual(path, other, 'extrusionOk')
      || path.commands.length !== other.commands.length
    ) return false;
    return path.commands.every((command, commandIndex) =>
      commandsEqual(command, other.commands[commandIndex]));
  });
}

function normalizePath(value: unknown, context: string): Readonly<CustomGeometryPath> {
  const path = readObject(value, PATH_KEYS, new Set(['width', 'height', 'commands']), context);
  const width = readSafeInteger(path.width, `${context} width`, true);
  const height = readSafeInteger(path.height, `${context} height`, true);
  const commands = readArray(path.commands, `${context} commands`).map((command, index) =>
    normalizeCommand(command, `${context} command ${index}`));
  if (commands[0] && commands[0].kind !== 'moveTo') {
    throw new TypeError(`${context} first command must be moveTo`);
  }

  const result: {
    width: number;
    height: number;
    fill?: CustomGeometryPathFill;
    stroke?: boolean;
    extrusionOk?: boolean;
    commands: readonly Readonly<CustomGeometryCommand>[];
  } = { width, height, commands: Object.freeze(commands) };
  if (Object.hasOwn(path, 'fill')) {
    if (typeof path.fill !== 'string' || !PATH_FILLS.has(path.fill as CustomGeometryPathFill)) {
      throw new TypeError(`${context} fill must be a supported path fill`);
    }
    result.fill = path.fill as CustomGeometryPathFill;
  }
  if (Object.hasOwn(path, 'stroke')) {
    if (typeof path.stroke !== 'boolean') {
      throw new TypeError(`${context} stroke must be boolean`);
    }
    result.stroke = path.stroke;
  }
  if (Object.hasOwn(path, 'extrusionOk')) {
    if (typeof path.extrusionOk !== 'boolean') {
      throw new TypeError(`${context} extrusionOk must be boolean`);
    }
    result.extrusionOk = path.extrusionOk;
  }
  return Object.freeze(result);
}

function normalizeCommand(value: unknown, context: string): Readonly<CustomGeometryCommand> {
  const candidate = readObjectData(value, context);
  const kind = candidate.kind;
  if (typeof kind !== 'string' || !COMMAND_KEYS.has(kind)) {
    throw new TypeError(`${context} kind must be a supported custom geometry command`);
  }
  requireKeys(candidate, COMMAND_KEYS.get(kind)!, COMMAND_KEYS.get(kind)!, context);
  switch (kind) {
    case 'moveTo':
    case 'lineTo':
      return Object.freeze({
        kind,
        point: normalizePoint(candidate.point, `${context} point`),
      });
    case 'arcTo':
      return Object.freeze({
        kind,
        widthRadius: readSafeInteger(candidate.widthRadius, `${context} widthRadius`, true),
        heightRadius: readSafeInteger(candidate.heightRadius, `${context} heightRadius`, true),
        startAngle: readSafeInteger(candidate.startAngle, `${context} startAngle`, false),
        sweepAngle: readSafeInteger(candidate.sweepAngle, `${context} sweepAngle`, false),
      });
    case 'quadraticBezierTo':
      return Object.freeze({
        kind,
        control: normalizePoint(candidate.control, `${context} control`),
        end: normalizePoint(candidate.end, `${context} end`),
      });
    case 'cubicBezierTo':
      return Object.freeze({
        kind,
        control1: normalizePoint(candidate.control1, `${context} control1`),
        control2: normalizePoint(candidate.control2, `${context} control2`),
        end: normalizePoint(candidate.end, `${context} end`),
      });
    case 'close':
      return Object.freeze({ kind });
    default:
      throw new TypeError(`${context} kind must be supported`);
  }
}

function normalizePoint(value: unknown, context: string): Readonly<CustomGeometryPoint> {
  const point = readObject(value, POINT_KEYS, POINT_KEYS, context);
  return Object.freeze({
    x: readSafeInteger(point.x, `${context} x`, false),
    y: readSafeInteger(point.y, `${context} y`, false),
  });
}

function renderPath(path: Readonly<CustomGeometryPath>, prefix: string): string {
  const attributes = [
    `w="${path.width}"`,
    `h="${path.height}"`,
    ...(Object.hasOwn(path, 'fill') ? [`fill="${path.fill}"`] : []),
    ...(Object.hasOwn(path, 'stroke') ? [`stroke="${path.stroke ? 1 : 0}"`] : []),
    ...(Object.hasOwn(path, 'extrusionOk')
      ? [`extrusionOk="${path.extrusionOk ? 1 : 0}"`]
      : []),
  ].join(' ');
  const commands = path.commands.map((command) => renderCommand(command, prefix)).join('');
  return `<${prefix}path ${attributes}>${commands}</${prefix}path>`;
}

function renderCommand(command: Readonly<CustomGeometryCommand>, prefix: string): string {
  switch (command.kind) {
    case 'moveTo':
      return `<${prefix}moveTo>${renderPoint(command.point, prefix)}</${prefix}moveTo>`;
    case 'lineTo':
      return `<${prefix}lnTo>${renderPoint(command.point, prefix)}</${prefix}lnTo>`;
    case 'arcTo':
      return `<${prefix}arcTo wR="${command.widthRadius}" hR="${command.heightRadius}" ` +
        `stAng="${command.startAngle}" swAng="${command.sweepAngle}"/>`;
    case 'quadraticBezierTo':
      return `<${prefix}quadBezTo>${renderPoint(command.control, prefix)}` +
        `${renderPoint(command.end, prefix)}</${prefix}quadBezTo>`;
    case 'cubicBezierTo':
      return `<${prefix}cubicBezTo>${renderPoint(command.control1, prefix)}` +
        `${renderPoint(command.control2, prefix)}${renderPoint(command.end, prefix)}` +
        `</${prefix}cubicBezTo>`;
    case 'close':
      return `<${prefix}close/>`;
  }
}

function renderPoint(point: Readonly<CustomGeometryPoint>, prefix: string): string {
  return `<${prefix}pt x="${point.x}" y="${point.y}"/>`;
}

function inspectCustomGeometryOwner(shape: XmlElement): CustomGeometryOwnerState | undefined {
  const geometry = resolveCustomGeometryElement(shape);
  if (!geometry) return undefined;
  const snapshot = parseCustomGeometryElement(geometry);
  if (!snapshot) return undefined;
  return {
    geometry,
    prefix: lexicalPrefix(geometry.name),
    snapshot,
  };
}

function resolveCustomGeometryElement(shape: XmlElement): XmlElement | undefined {
  if (shape.localName !== 'sp' || namespaceUri(shape) !== PRESENTATION_NAMESPACE) {
    return undefined;
  }
  const properties = directChildren(shape).filter(({ localName }) => localName === 'spPr');
  if (properties.length !== 1 || namespaceUri(properties[0]!) !== PRESENTATION_NAMESPACE) {
    return undefined;
  }
  const geometries = directChildren(properties[0]!).filter(
    ({ localName }) => localName === 'prstGeom' || localName === 'custGeom',
  );
  const geometry = geometries[0];
  if (
    geometries.length !== 1
    || !geometry
    || geometry.localName !== 'custGeom'
    || namespaceUri(geometry) !== DRAWING_NAMESPACE
  ) return undefined;
  return geometry;
}

function parseCustomGeometryElement(geometry: XmlElement): NormalizedCustomGeometry | undefined {
  if (nonNamespaceAttributes(geometry).length !== 0 || hasNonWhitespaceText(geometry)) {
    return undefined;
  }
  const children = directChildren(geometry);
  let previousStage = -1;
  for (const child of children) {
    if (namespaceUri(child) !== DRAWING_NAMESPACE) return undefined;
    const stage = CUSTOM_CHILD_STAGES.get(child.localName);
    if (stage === undefined || stage < previousStage) return undefined;
    previousStage = stage;
  }

  for (const name of EMPTY_LIST_NAMES) {
    const lists = children.filter(({ localName }) => localName === name);
    if (
      lists.length > 1
      || (lists[0] && !isEmptyElement(lists[0]))
    ) return undefined;
  }

  const rectangles = children.filter(({ localName }) => localName === 'rect');
  if (rectangles.length > 1 || (rectangles[0] && !isDefaultRectangle(rectangles[0]))) {
    return undefined;
  }

  const pathLists = children.filter(({ localName }) => localName === 'pathLst');
  const pathList = pathLists[0];
  if (
    pathLists.length !== 1
    || !pathList
    || nonNamespaceAttributes(pathList).length !== 0
    || hasNonWhitespaceText(pathList)
  ) return undefined;
  const pathElements = directChildren(pathList);
  if (
    pathElements.length === 0
    || pathElements.some(
      (path) => path.localName !== 'path' || namespaceUri(path) !== DRAWING_NAMESPACE,
    )
  ) return undefined;

  const paths: CustomGeometryPath[] = [];
  for (const pathElement of pathElements) {
    const path = parsePath(pathElement);
    if (!path) return undefined;
    paths.push(path);
  }
  try {
    return normalizeCustomGeometry({ paths }, 'Custom geometry');
  } catch {
    return undefined;
  }
}

function parsePath(path: XmlElement): CustomGeometryPath | undefined {
  if (hasNonWhitespaceText(path)) return undefined;
  const attributes = readXmlAttributes(
    path,
    new Set(['w', 'h', 'fill', 'stroke', 'extrusionOk']),
    new Set(['w', 'h']),
  );
  if (!attributes) return undefined;
  const width = parseInteger(attributes.w, true);
  const height = parseInteger(attributes.h, true);
  if (width === undefined || height === undefined) return undefined;

  const commands: CustomGeometryCommand[] = [];
  for (const commandElement of directChildren(path)) {
    if (namespaceUri(commandElement) !== DRAWING_NAMESPACE) return undefined;
    const command = parseCommandElement(commandElement);
    if (!command) return undefined;
    commands.push(command);
  }
  if (commands[0] && commands[0].kind !== 'moveTo') return undefined;

  const result: {
    width: number;
    height: number;
    fill?: CustomGeometryPathFill;
    stroke?: boolean;
    extrusionOk?: boolean;
    commands: CustomGeometryCommand[];
  } = { width, height, commands };
  if (Object.hasOwn(attributes, 'fill')) {
    if (!PATH_FILLS.has(attributes.fill as CustomGeometryPathFill)) return undefined;
    result.fill = attributes.fill as CustomGeometryPathFill;
  }
  if (Object.hasOwn(attributes, 'stroke')) {
    const stroke = parseBoolean(attributes.stroke);
    if (stroke === undefined) return undefined;
    result.stroke = stroke;
  }
  if (Object.hasOwn(attributes, 'extrusionOk')) {
    const extrusionOk = parseBoolean(attributes.extrusionOk);
    if (extrusionOk === undefined) return undefined;
    result.extrusionOk = extrusionOk;
  }
  return result;
}

function parseCommandElement(element: XmlElement): CustomGeometryCommand | undefined {
  if (hasNonWhitespaceText(element)) return undefined;
  switch (element.localName) {
    case 'moveTo':
    case 'lnTo': {
      if (nonNamespaceAttributes(element).length !== 0) return undefined;
      const points = directChildren(element);
      if (points.length !== 1) return undefined;
      const point = parsePointElement(points[0]!);
      if (!point) return undefined;
      return element.localName === 'moveTo'
        ? { kind: 'moveTo', point }
        : { kind: 'lineTo', point };
    }
    case 'arcTo': {
      if (directChildren(element).length !== 0) return undefined;
      const attributes = readXmlAttributes(
        element,
        new Set(['wR', 'hR', 'stAng', 'swAng']),
        new Set(['wR', 'hR', 'stAng', 'swAng']),
      );
      if (!attributes) return undefined;
      const widthRadius = parseInteger(attributes.wR, true);
      const heightRadius = parseInteger(attributes.hR, true);
      const startAngle = parseInteger(attributes.stAng, false);
      const sweepAngle = parseInteger(attributes.swAng, false);
      if (
        widthRadius === undefined
        || heightRadius === undefined
        || startAngle === undefined
        || sweepAngle === undefined
      ) return undefined;
      return { kind: 'arcTo', widthRadius, heightRadius, startAngle, sweepAngle };
    }
    case 'quadBezTo': {
      if (nonNamespaceAttributes(element).length !== 0) return undefined;
      const points = directChildren(element);
      if (points.length !== 2) return undefined;
      const control = parsePointElement(points[0]!);
      const end = parsePointElement(points[1]!);
      return control && end ? { kind: 'quadraticBezierTo', control, end } : undefined;
    }
    case 'cubicBezTo': {
      if (nonNamespaceAttributes(element).length !== 0) return undefined;
      const points = directChildren(element);
      if (points.length !== 3) return undefined;
      const control1 = parsePointElement(points[0]!);
      const control2 = parsePointElement(points[1]!);
      const end = parsePointElement(points[2]!);
      return control1 && control2 && end
        ? { kind: 'cubicBezierTo', control1, control2, end }
        : undefined;
    }
    case 'close':
      return nonNamespaceAttributes(element).length === 0
        && directChildren(element).length === 0
        ? { kind: 'close' }
        : undefined;
    default:
      return undefined;
  }
}

function parsePointElement(element: XmlElement): CustomGeometryPoint | undefined {
  if (
    element.localName !== 'pt'
    || namespaceUri(element) !== DRAWING_NAMESPACE
    || directChildren(element).length !== 0
    || hasNonWhitespaceText(element)
  ) return undefined;
  const attributes = readXmlAttributes(element, POINT_KEYS, POINT_KEYS);
  if (!attributes) return undefined;
  const x = parseInteger(attributes.x, false);
  const y = parseInteger(attributes.y, false);
  return x === undefined || y === undefined ? undefined : { x, y };
}

function isEmptyElement(element: XmlElement): boolean {
  return nonNamespaceAttributes(element).length === 0
    && directChildren(element).length === 0
    && !hasNonWhitespaceText(element);
}

function isDefaultRectangle(element: XmlElement): boolean {
  if (directChildren(element).length !== 0 || hasNonWhitespaceText(element)) return false;
  const attributes = readXmlAttributes(
    element,
    new Set(['l', 't', 'r', 'b']),
    new Set(['l', 't', 'r', 'b']),
  );
  return attributes !== undefined
    && attributes.l === 'l'
    && attributes.t === 't'
    && attributes.r === 'r'
    && attributes.b === 'b';
}

function readObject(
  value: unknown,
  allowed: ReadonlySet<string>,
  required: ReadonlySet<string>,
  context: string,
): Record<string, unknown> {
  const result = readObjectData(value, context);
  requireKeys(result, allowed, required, context);
  return result;
}

function readObjectData(value: unknown, context: string): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${context} must be an object`);
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError(`${context} must be an ordinary object`);
  }
  const result = Object.create(null) as Record<string, unknown>;
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string') {
      throw new TypeError(`${context} contains unsupported property ${String(key)}`);
    }
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`${context} property ${key} must be a data property`);
    }
    result[key] = descriptor.value;
  }
  return result;
}

function requireKeys(
  value: Record<string, unknown>,
  allowed: ReadonlySet<string>,
  required: ReadonlySet<string>,
  context: string,
): void {
  for (const key of Object.keys(value)) {
    if (!allowed.has(key)) {
      throw new TypeError(`${context} contains unsupported property ${key}`);
    }
  }
  for (const key of required) {
    if (!Object.hasOwn(value, key)) {
      throw new TypeError(`${context} is missing property ${key}`);
    }
  }
}

function readArray(value: unknown, context: string): readonly unknown[] {
  if (!Array.isArray(value) || Object.getPrototypeOf(value) !== Array.prototype) {
    throw new TypeError(`${context} must be an ordinary array`);
  }
  const keys = Reflect.ownKeys(value);
  if (keys.length !== value.length + 1 || !keys.includes('length')) {
    throw new TypeError(`${context} must be a dense array without extra properties`);
  }
  const result: unknown[] = [];
  for (let index = 0; index < value.length; index += 1) {
    const key = String(index);
    if (!keys.includes(key)) throw new TypeError(`${context} must be a dense array`);
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`${context} item ${index} must be a data property`);
    }
    result.push(descriptor.value);
  }
  return result;
}

function readSafeInteger(value: unknown, context: string, positive: boolean): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  if (!Number.isSafeInteger(value)) throw new RangeError(`${context} must be a safe integer`);
  if (positive && value <= 0) throw new RangeError(`${context} must be positive`);
  return normalizeNegativeZero(value);
}

function readXmlAttributes(
  element: XmlElement,
  allowed: ReadonlySet<string>,
  required: ReadonlySet<string>,
): Record<string, string> | undefined {
  const result: Record<string, string> = {};
  for (const attribute of nonNamespaceAttributes(element)) {
    if (!allowed.has(attribute.name) || Object.hasOwn(result, attribute.name)) return undefined;
    result[attribute.name] = attribute.value;
  }
  for (const name of required) if (!Object.hasOwn(result, name)) return undefined;
  return result;
}

function parseInteger(value: string | undefined, positive: boolean): number | undefined {
  if (!value || !INTEGER_PATTERN.test(value)) return undefined;
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed) || (positive && parsed <= 0)) return undefined;
  return normalizeNegativeZero(parsed);
}

function parseBoolean(value: string | undefined): boolean | undefined {
  if (value === '1' || value === 'true' || value === 'on') return true;
  if (value === '0' || value === 'false' || value === 'off') return false;
  return undefined;
}

function optionalPropertyEqual<T extends object>(
  left: T,
  right: T,
  key: keyof T,
): boolean {
  return Object.hasOwn(left, key) === Object.hasOwn(right, key)
    && left[key] === right[key];
}

function commandsEqual(
  left: Readonly<CustomGeometryCommand>,
  right: Readonly<CustomGeometryCommand> | undefined,
): boolean {
  if (!right || left.kind !== right.kind) return false;
  switch (left.kind) {
    case 'moveTo':
    case 'lineTo':
      return right.kind === left.kind && pointsEqual(left.point, right.point);
    case 'arcTo':
      return right.kind === 'arcTo'
        && left.widthRadius === right.widthRadius
        && left.heightRadius === right.heightRadius
        && left.startAngle === right.startAngle
        && left.sweepAngle === right.sweepAngle;
    case 'quadraticBezierTo':
      return right.kind === 'quadraticBezierTo'
        && pointsEqual(left.control, right.control)
        && pointsEqual(left.end, right.end);
    case 'cubicBezierTo':
      return right.kind === 'cubicBezierTo'
        && pointsEqual(left.control1, right.control1)
        && pointsEqual(left.control2, right.control2)
        && pointsEqual(left.end, right.end);
    case 'close':
      return right.kind === 'close';
  }
}

function pointsEqual(
  left: Readonly<CustomGeometryPoint>,
  right: Readonly<CustomGeometryPoint>,
): boolean {
  return left.x === right.x && left.y === right.y;
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

function normalizeNegativeZero(value: number): number {
  return Object.is(value, -0) ? 0 : value;
}
