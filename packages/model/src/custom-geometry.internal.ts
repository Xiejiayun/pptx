import {
  escapeXmlAttribute,
  type LosslessXmlDocument,
  type XmlElement,
} from '@pptx/lossless-xml';
import type {
  CustomGeometry,
  CustomGeometryCommand,
  CustomGeometryConnectionSite,
  CustomGeometryFormula,
  CustomGeometryGuide,
  CustomGeometryHandle,
  CustomGeometryPath,
  CustomGeometryPathFill,
  CustomGeometryPoint,
  CustomGeometryPolarHandle,
  CustomGeometryTextRectangle,
  CustomGeometryValue,
  CustomGeometryXyHandle,
} from './custom-geometry.js';
import { ModelParseError } from './errors.js';
import { PRESET_SHAPE_TYPES } from './preset-shape.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const ROOT_KEYS = new Set([
  'adjustments',
  'guides',
  'handles',
  'connectionSites',
  'textRectangle',
  'paths',
]);
const ROOT_REQUIRED_KEYS = new Set(['paths']);
const GUIDE_KEYS = new Set(['name', 'formula']);
const FORMULA_KEYS = new Set(['operator', 'operands']);
const XY_HANDLE_KEYS = new Set([
  'kind',
  'position',
  'xGuide',
  'minX',
  'maxX',
  'yGuide',
  'minY',
  'maxY',
]);
const POLAR_HANDLE_KEYS = new Set([
  'kind',
  'position',
  'radiusGuide',
  'minRadius',
  'maxRadius',
  'angleGuide',
  'minAngle',
  'maxAngle',
]);
const HANDLE_REQUIRED_KEYS = new Set(['kind', 'position']);
const CONNECTION_SITE_KEYS = new Set(['position', 'angle']);
const TEXT_RECTANGLE_KEYS = new Set(['left', 'top', 'right', 'bottom']);
const DEFAULT_TEXT_RECTANGLE: Readonly<CustomGeometryTextRectangle> = Object.freeze({
  left: 'l',
  top: 't',
  right: 'r',
  bottom: 'b',
});
const XY_HANDLE_ATTRIBUTE_KEYS = new Set([
  'gdRefX',
  'minX',
  'maxX',
  'gdRefY',
  'minY',
  'maxY',
]);
const POLAR_HANDLE_ATTRIBUTE_KEYS = new Set([
  'gdRefR',
  'minR',
  'maxR',
  'gdRefAng',
  'minAng',
  'maxAng',
]);
const NO_REQUIRED_KEYS = new Set<string>();
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
const INTEGER_PATTERN = /^[+-]?\d+$/;
const XML_WHITESPACE_PATTERN = /[ \t\r\n]/;
const FORMULA_ARITIES: ReadonlyMap<string, number> = new Map([
  ['val', 1],
  ['abs', 1],
  ['sqrt', 1],
  ['at2', 2],
  ['cos', 2],
  ['max', 2],
  ['min', 2],
  ['sin', 2],
  ['tan', 2],
  ['*/', 3],
  ['+-', 3],
  ['+/', 3],
  ['?:', 3],
  ['cat2', 3],
  ['mod', 3],
  ['pin', 3],
  ['sat2', 3],
]);
const PRESET_SHAPE_TYPE_SET: ReadonlySet<string> = new Set(PRESET_SHAPE_TYPES);

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
  const root = readObject(value, ROOT_KEYS, ROOT_REQUIRED_KEYS, context);
  const guideNames = new Set<string>();
  const adjustments = Object.hasOwn(root, 'adjustments')
    ? normalizeGuideList(root.adjustments, `${context} adjustments`, guideNames)
    : undefined;
  const guides = Object.hasOwn(root, 'guides')
    ? normalizeGuideList(root.guides, `${context} guides`, guideNames)
    : undefined;
  const handles = Object.hasOwn(root, 'handles')
    ? normalizeHandleList(root.handles, `${context} handles`)
    : undefined;
  const connectionSites = Object.hasOwn(root, 'connectionSites')
    ? normalizeConnectionSiteList(root.connectionSites, `${context} connectionSites`)
    : undefined;
  const textRectangle = Object.hasOwn(root, 'textRectangle')
    ? normalizeTextRectangle(root.textRectangle, `${context} textRectangle`)
    : undefined;
  const paths = readArray(root.paths, `${context} paths`);
  if (paths.length === 0) throw new RangeError(`${context} paths must not be empty`);
  const normalizedPaths = paths.map((path, index) =>
    normalizePath(path, `${context} path ${index}`));
  return Object.freeze({
    ...(adjustments?.length ? { adjustments } : {}),
    ...(guides?.length ? { guides } : {}),
    ...(handles?.length ? { handles } : {}),
    ...(connectionSites?.length ? { connectionSites } : {}),
    ...(textRectangle ? { textRectangle } : {}),
    paths: Object.freeze(normalizedPaths),
  });
}

export function renderCustomGeometry(
  geometry: NormalizedCustomGeometry,
  prefix: string,
): string {
  const paths = geometry.paths.map((path) => renderPath(path, prefix)).join('');
  return `<${prefix}custGeom>${renderGuideList('avLst', geometry.adjustments, prefix)}` +
    `${renderGuideList('gdLst', geometry.guides, prefix)}` +
    `${renderHandleList(geometry.handles, prefix)}` +
    `${renderConnectionSiteList(geometry.connectionSites, prefix)}` +
    `${renderTextRectangle(geometry.textRectangle, prefix)}` +
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
  const customState = inspectCustomGeometryOwner(shape);
  const state = customState ?? inspectPresetGeometryOwner(shape);
  if (!state) {
    throw new ModelParseError('Shape custom geometry is not safely editable', partUri);
  }
  if (customState && customGeometryEqual(customState.snapshot, geometry)) return false;

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
  if (
    !guideListsEqual(left.adjustments, right.adjustments)
    || !guideListsEqual(left.guides, right.guides)
    || !handleListsEqual(left.handles, right.handles)
    || !connectionSiteListsEqual(left.connectionSites, right.connectionSites)
    || !textRectanglesEqual(left.textRectangle, right.textRectangle)
  ) return false;
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

function normalizeGuideList(
  value: unknown,
  context: string,
  names: Set<string>,
): readonly Readonly<CustomGeometryGuide>[] {
  const guides = readArray(value, context).map((item, index) => {
    const itemContext = `${context} item ${index}`;
    const guide = readObject(item, GUIDE_KEYS, GUIDE_KEYS, itemContext);
    const name = normalizeCustomGeometryToken(guide.name, `${itemContext} name`);
    if (names.has(name)) throw new TypeError(`${itemContext} name must be unique`);
    names.add(name);
    return Object.freeze({
      name,
      formula: normalizeFormula(guide.formula, `${itemContext} formula`),
    });
  });
  return Object.freeze(guides);
}

function normalizeFormula(
  value: unknown,
  context: string,
): Readonly<CustomGeometryFormula> {
  const formula = readObject(value, FORMULA_KEYS, FORMULA_KEYS, context);
  if (typeof formula.operator !== 'string') {
    throw new TypeError(`${context} operator must be a supported string`);
  }
  const arity = FORMULA_ARITIES.get(formula.operator);
  if (arity === undefined) throw new TypeError(`${context} operator must be supported`);
  const operands = readArray(formula.operands, `${context} operands`);
  if (operands.length !== arity) {
    throw new RangeError(`${context} operands must contain exactly ${arity} items`);
  }
  return Object.freeze({
    operator: formula.operator,
    operands: Object.freeze(operands.map((operand, index) =>
      normalizeCustomGeometryValue(operand, `${context} operand ${index}`, false))),
  }) as Readonly<CustomGeometryFormula>;
}

function normalizeHandleList(
  value: unknown,
  context: string,
): readonly Readonly<CustomGeometryHandle>[] {
  const handles = readArray(value, context).map((item, index) =>
    normalizeHandle(item, `${context} item ${index}`));
  return Object.freeze(handles);
}

function normalizeHandle(value: unknown, context: string): Readonly<CustomGeometryHandle> {
  const handle = readObjectData(value, context);
  const kind = handle.kind;
  if (kind !== 'xy' && kind !== 'polar') {
    throw new TypeError(`${context} kind must be xy or polar`);
  }
  requireKeys(
    handle,
    kind === 'xy' ? XY_HANDLE_KEYS : POLAR_HANDLE_KEYS,
    HANDLE_REQUIRED_KEYS,
    context,
  );
  const position = normalizePoint(handle.position, `${context} position`);
  if (kind === 'xy') {
    const result: {
      kind: 'xy';
      position: Readonly<CustomGeometryPoint>;
      xGuide?: string;
      minX?: CustomGeometryValue;
      maxX?: CustomGeometryValue;
      yGuide?: string;
      minY?: CustomGeometryValue;
      maxY?: CustomGeometryValue;
    } = { kind, position };
    if (Object.hasOwn(handle, 'xGuide')) {
      result.xGuide = normalizeCustomGeometryToken(handle.xGuide, `${context} xGuide`);
    }
    if (Object.hasOwn(handle, 'minX')) {
      result.minX = normalizeCustomGeometryValue(handle.minX, `${context} minX`, false);
    }
    if (Object.hasOwn(handle, 'maxX')) {
      result.maxX = normalizeCustomGeometryValue(handle.maxX, `${context} maxX`, false);
    }
    if (Object.hasOwn(handle, 'yGuide')) {
      result.yGuide = normalizeCustomGeometryToken(handle.yGuide, `${context} yGuide`);
    }
    if (Object.hasOwn(handle, 'minY')) {
      result.minY = normalizeCustomGeometryValue(handle.minY, `${context} minY`, false);
    }
    if (Object.hasOwn(handle, 'maxY')) {
      result.maxY = normalizeCustomGeometryValue(handle.maxY, `${context} maxY`, false);
    }
    return Object.freeze(result);
  }

  const result: {
    kind: 'polar';
    position: Readonly<CustomGeometryPoint>;
    radiusGuide?: string;
    minRadius?: CustomGeometryValue;
    maxRadius?: CustomGeometryValue;
    angleGuide?: string;
    minAngle?: CustomGeometryValue;
    maxAngle?: CustomGeometryValue;
  } = { kind, position };
  if (Object.hasOwn(handle, 'radiusGuide')) {
    result.radiusGuide = normalizeCustomGeometryToken(
      handle.radiusGuide,
      `${context} radiusGuide`,
    );
  }
  if (Object.hasOwn(handle, 'minRadius')) {
    result.minRadius = normalizeCustomGeometryValue(
      handle.minRadius,
      `${context} minRadius`,
      false,
    );
  }
  if (Object.hasOwn(handle, 'maxRadius')) {
    result.maxRadius = normalizeCustomGeometryValue(
      handle.maxRadius,
      `${context} maxRadius`,
      false,
    );
  }
  if (Object.hasOwn(handle, 'angleGuide')) {
    result.angleGuide = normalizeCustomGeometryToken(
      handle.angleGuide,
      `${context} angleGuide`,
    );
  }
  if (Object.hasOwn(handle, 'minAngle')) {
    result.minAngle = normalizeCustomGeometryValue(
      handle.minAngle,
      `${context} minAngle`,
      false,
    );
  }
  if (Object.hasOwn(handle, 'maxAngle')) {
    result.maxAngle = normalizeCustomGeometryValue(
      handle.maxAngle,
      `${context} maxAngle`,
      false,
    );
  }
  return Object.freeze(result);
}

function normalizeConnectionSiteList(
  value: unknown,
  context: string,
): readonly Readonly<CustomGeometryConnectionSite>[] {
  const sites = readArray(value, context).map((item, index) => {
    const itemContext = `${context} item ${index}`;
    const site = readObject(item, CONNECTION_SITE_KEYS, CONNECTION_SITE_KEYS, itemContext);
    return Object.freeze({
      position: normalizePoint(site.position, `${itemContext} position`),
      angle: normalizeCustomGeometryValue(site.angle, `${itemContext} angle`, false),
    });
  });
  return Object.freeze(sites);
}

function normalizeTextRectangle(
  value: unknown,
  context: string,
): Readonly<CustomGeometryTextRectangle> | undefined {
  const rectangle = readObject(value, TEXT_RECTANGLE_KEYS, TEXT_RECTANGLE_KEYS, context);
  const normalized = Object.freeze({
    left: normalizeCustomGeometryValue(rectangle.left, `${context} left`, false),
    top: normalizeCustomGeometryValue(rectangle.top, `${context} top`, false),
    right: normalizeCustomGeometryValue(rectangle.right, `${context} right`, false),
    bottom: normalizeCustomGeometryValue(rectangle.bottom, `${context} bottom`, false),
  });
  return isDefaultTextRectangle(normalized) ? undefined : normalized;
}

function isDefaultTextRectangle(
  rectangle: Readonly<CustomGeometryTextRectangle>,
): boolean {
  return rectangle.left === DEFAULT_TEXT_RECTANGLE.left
    && rectangle.top === DEFAULT_TEXT_RECTANGLE.top
    && rectangle.right === DEFAULT_TEXT_RECTANGLE.right
    && rectangle.bottom === DEFAULT_TEXT_RECTANGLE.bottom;
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
        widthRadius: normalizeCustomGeometryValue(
          candidate.widthRadius,
          `${context} widthRadius`,
          true,
        ),
        heightRadius: normalizeCustomGeometryValue(
          candidate.heightRadius,
          `${context} heightRadius`,
          true,
        ),
        startAngle: normalizeCustomGeometryValue(
          candidate.startAngle,
          `${context} startAngle`,
          false,
        ),
        sweepAngle: normalizeCustomGeometryValue(
          candidate.sweepAngle,
          `${context} sweepAngle`,
          false,
        ),
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
    x: normalizeCustomGeometryValue(point.x, `${context} x`, false),
    y: normalizeCustomGeometryValue(point.y, `${context} y`, false),
  });
}

function normalizeCustomGeometryValue(
  value: unknown,
  context: string,
  positive: boolean,
): CustomGeometryValue {
  if (typeof value === 'number') return readSafeInteger(value, context, positive);
  return normalizeCustomGeometryToken(value, context);
}

function normalizeCustomGeometryToken(value: unknown, context: string): string {
  if (typeof value !== 'string') throw new TypeError(`${context} must be a number or string`);
  if (
    value.length === 0
    || XML_WHITESPACE_PATTERN.test(value)
    || INTEGER_PATTERN.test(value)
    || containsInvalidXmlCharacter(value)
  ) {
    throw new TypeError(`${context} must be a non-decimal XML-safe token`);
  }
  return value;
}

function renderGuideList(
  name: 'avLst' | 'gdLst',
  guides: readonly Readonly<CustomGeometryGuide>[] | undefined,
  prefix: string,
): string {
  if (!guides?.length) return `<${prefix}${name}/>`;
  const children = guides.map((guide) =>
    `<${prefix}gd name="${escapeXmlAttribute(guide.name)}" ` +
    `fmla="${escapeXmlAttribute(renderFormula(guide.formula))}"/>`).join('');
  return `<${prefix}${name}>${children}</${prefix}${name}>`;
}

function renderFormula(formula: Readonly<CustomGeometryFormula>): string {
  return [formula.operator, ...formula.operands.map(String)].join(' ');
}

function renderHandleList(
  handles: readonly Readonly<CustomGeometryHandle>[] | undefined,
  prefix: string,
): string {
  if (!handles?.length) return `<${prefix}ahLst/>`;
  const children = handles.map((handle) => renderHandle(handle, prefix)).join('');
  return `<${prefix}ahLst>${children}</${prefix}ahLst>`;
}

function renderHandle(handle: Readonly<CustomGeometryHandle>, prefix: string): string {
  if (handle.kind === 'xy') {
    const attributes = [
      renderHandleAttribute(handle, 'xGuide', 'gdRefX', handle.xGuide),
      renderHandleAttribute(handle, 'minX', 'minX', handle.minX),
      renderHandleAttribute(handle, 'maxX', 'maxX', handle.maxX),
      renderHandleAttribute(handle, 'yGuide', 'gdRefY', handle.yGuide),
      renderHandleAttribute(handle, 'minY', 'minY', handle.minY),
      renderHandleAttribute(handle, 'maxY', 'maxY', handle.maxY),
    ].filter((value) => value !== undefined).join(' ');
    return `<${prefix}ahXY${attributes ? ` ${attributes}` : ''}>` +
      `${renderPoint(handle.position, prefix, 'pos')}</${prefix}ahXY>`;
  }
  const attributes = [
    renderHandleAttribute(handle, 'radiusGuide', 'gdRefR', handle.radiusGuide),
    renderHandleAttribute(handle, 'minRadius', 'minR', handle.minRadius),
    renderHandleAttribute(handle, 'maxRadius', 'maxR', handle.maxRadius),
    renderHandleAttribute(handle, 'angleGuide', 'gdRefAng', handle.angleGuide),
    renderHandleAttribute(handle, 'minAngle', 'minAng', handle.minAngle),
    renderHandleAttribute(handle, 'maxAngle', 'maxAng', handle.maxAngle),
  ].filter((value) => value !== undefined).join(' ');
  return `<${prefix}ahPolar${attributes ? ` ${attributes}` : ''}>` +
    `${renderPoint(handle.position, prefix, 'pos')}</${prefix}ahPolar>`;
}

function renderHandleAttribute(
  handle: Readonly<CustomGeometryHandle>,
  property: string,
  attribute: string,
  value: CustomGeometryValue | undefined,
): string | undefined {
  return Object.hasOwn(handle, property)
    ? `${attribute}="${renderCustomGeometryValue(value!)}"`
    : undefined;
}

function renderConnectionSiteList(
  sites: readonly Readonly<CustomGeometryConnectionSite>[] | undefined,
  prefix: string,
): string {
  if (!sites?.length) return `<${prefix}cxnLst/>`;
  const children = sites.map((site) =>
    `<${prefix}cxn ang="${renderCustomGeometryValue(site.angle)}">` +
    `${renderPoint(site.position, prefix, 'pos')}</${prefix}cxn>`).join('');
  return `<${prefix}cxnLst>${children}</${prefix}cxnLst>`;
}

function renderTextRectangle(
  rectangle: Readonly<CustomGeometryTextRectangle> | undefined,
  prefix: string,
): string {
  const value = rectangle ?? DEFAULT_TEXT_RECTANGLE;
  return `<${prefix}rect l="${renderCustomGeometryValue(value.left)}" ` +
    `t="${renderCustomGeometryValue(value.top)}" ` +
    `r="${renderCustomGeometryValue(value.right)}" ` +
    `b="${renderCustomGeometryValue(value.bottom)}"/>`;
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
      return `<${prefix}arcTo wR="${renderCustomGeometryValue(command.widthRadius)}" ` +
        `hR="${renderCustomGeometryValue(command.heightRadius)}" ` +
        `stAng="${renderCustomGeometryValue(command.startAngle)}" ` +
        `swAng="${renderCustomGeometryValue(command.sweepAngle)}"/>`;
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

function renderPoint(
  point: Readonly<CustomGeometryPoint>,
  prefix: string,
  elementName = 'pt',
): string {
  return `<${prefix}${elementName} x="${renderCustomGeometryValue(point.x)}" ` +
    `y="${renderCustomGeometryValue(point.y)}"/>`;
}

function renderCustomGeometryValue(value: CustomGeometryValue): string {
  return escapeXmlAttribute(String(value));
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

function inspectPresetGeometryOwner(
  shape: XmlElement,
): Omit<CustomGeometryOwnerState, 'snapshot'> | undefined {
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
  return { geometry, prefix: lexicalPrefix(geometry.name) };
}

function resolveCustomGeometryElement(shape: XmlElement): XmlElement | undefined {
  const geometry = resolveGeometryChoice(shape);
  return geometry?.localName === 'custGeom' ? geometry : undefined;
}

function resolveGeometryChoice(shape: XmlElement): XmlElement | undefined {
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

  const adjustments = parseGuideList(children, 'avLst');
  const guides = parseGuideList(children, 'gdLst');
  const handles = parseHandleList(children);
  const connectionSites = parseConnectionSiteList(children);
  if (
    adjustments === undefined
    || guides === undefined
    || handles === undefined
    || connectionSites === undefined
  ) return undefined;

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
    return normalizeCustomGeometry({
      ...(adjustments.length ? { adjustments } : {}),
      ...(guides.length ? { guides } : {}),
      ...(handles.length ? { handles } : {}),
      ...(connectionSites.length ? { connectionSites } : {}),
      paths,
    }, 'Custom geometry');
  } catch {
    return undefined;
  }
}

function parseGuideList(
  children: readonly XmlElement[],
  name: 'avLst' | 'gdLst',
): CustomGeometryGuide[] | undefined {
  const lists = children.filter(({ localName }) => localName === name);
  if (lists.length > 1) return undefined;
  const list = lists[0];
  if (!list) return [];
  if (nonNamespaceAttributes(list).length !== 0 || hasNonWhitespaceText(list)) {
    return undefined;
  }

  const guides: CustomGeometryGuide[] = [];
  for (const element of directChildren(list)) {
    if (
      element.localName !== 'gd'
      || namespaceUri(element) !== DRAWING_NAMESPACE
      || directChildren(element).length !== 0
      || hasNonWhitespaceText(element)
    ) return undefined;
    const attributes = readXmlAttributes(
      element,
      new Set(['name', 'fmla']),
      new Set(['name', 'fmla']),
    );
    if (!attributes) return undefined;
    const guideName = parseCustomGeometryToken(attributes.name);
    const formula = parseFormulaAttribute(attributes.fmla);
    if (guideName === undefined || formula === undefined) return undefined;
    guides.push({ name: guideName, formula });
  }
  return guides;
}

function parseHandleList(children: readonly XmlElement[]): CustomGeometryHandle[] | undefined {
  const lists = children.filter(({ localName }) => localName === 'ahLst');
  if (lists.length > 1) return undefined;
  const list = lists[0];
  if (!list) return [];
  if (nonNamespaceAttributes(list).length !== 0 || hasNonWhitespaceText(list)) {
    return undefined;
  }

  const handles: CustomGeometryHandle[] = [];
  for (const element of directChildren(list)) {
    if (namespaceUri(element) !== DRAWING_NAMESPACE) return undefined;
    const handle = parseHandleElement(element);
    if (!handle) return undefined;
    handles.push(handle);
  }
  return handles;
}

function parseConnectionSiteList(
  children: readonly XmlElement[],
): CustomGeometryConnectionSite[] | undefined {
  const lists = children.filter(({ localName }) => localName === 'cxnLst');
  if (lists.length > 1) return undefined;
  const list = lists[0];
  if (!list) return [];
  if (nonNamespaceAttributes(list).length !== 0 || hasNonWhitespaceText(list)) {
    return undefined;
  }

  const sites: CustomGeometryConnectionSite[] = [];
  for (const element of directChildren(list)) {
    if (
      element.localName !== 'cxn'
      || namespaceUri(element) !== DRAWING_NAMESPACE
      || hasNonWhitespaceText(element)
    ) return undefined;
    const attributes = readXmlAttributes(element, new Set(['ang']), new Set(['ang']));
    const positions = directChildren(element);
    if (!attributes || positions.length !== 1) return undefined;
    const angle = parseCustomGeometryValue(attributes.ang, false);
    const position = parsePointElement(positions[0]!, 'pos');
    if (angle === undefined || !position) return undefined;
    sites.push({ position, angle });
  }
  return sites;
}

function parseHandleElement(element: XmlElement): CustomGeometryHandle | undefined {
  if (hasNonWhitespaceText(element)) return undefined;
  const positions = directChildren(element);
  if (positions.length !== 1) return undefined;
  const position = parsePointElement(positions[0]!, 'pos');
  if (!position) return undefined;
  switch (element.localName) {
    case 'ahXY':
      return parseXyHandle(element, position);
    case 'ahPolar':
      return parsePolarHandle(element, position);
    default:
      return undefined;
  }
}

function parseXyHandle(
  element: XmlElement,
  position: CustomGeometryPoint,
): CustomGeometryXyHandle | undefined {
  const attributes = readXmlAttributes(
    element,
    XY_HANDLE_ATTRIBUTE_KEYS,
    NO_REQUIRED_KEYS,
  );
  if (!attributes) return undefined;
  const result: {
    kind: 'xy';
    position: CustomGeometryPoint;
    xGuide?: string;
    minX?: CustomGeometryValue;
    maxX?: CustomGeometryValue;
    yGuide?: string;
    minY?: CustomGeometryValue;
    maxY?: CustomGeometryValue;
  } = { kind: 'xy', position };
  if (Object.hasOwn(attributes, 'gdRefX')) {
    const value = parseCustomGeometryToken(attributes.gdRefX);
    if (value === undefined) return undefined;
    result.xGuide = value;
  }
  if (Object.hasOwn(attributes, 'minX')) {
    const value = parseCustomGeometryValue(attributes.minX, false);
    if (value === undefined) return undefined;
    result.minX = value;
  }
  if (Object.hasOwn(attributes, 'maxX')) {
    const value = parseCustomGeometryValue(attributes.maxX, false);
    if (value === undefined) return undefined;
    result.maxX = value;
  }
  if (Object.hasOwn(attributes, 'gdRefY')) {
    const value = parseCustomGeometryToken(attributes.gdRefY);
    if (value === undefined) return undefined;
    result.yGuide = value;
  }
  if (Object.hasOwn(attributes, 'minY')) {
    const value = parseCustomGeometryValue(attributes.minY, false);
    if (value === undefined) return undefined;
    result.minY = value;
  }
  if (Object.hasOwn(attributes, 'maxY')) {
    const value = parseCustomGeometryValue(attributes.maxY, false);
    if (value === undefined) return undefined;
    result.maxY = value;
  }
  return result;
}

function parsePolarHandle(
  element: XmlElement,
  position: CustomGeometryPoint,
): CustomGeometryPolarHandle | undefined {
  const attributes = readXmlAttributes(
    element,
    POLAR_HANDLE_ATTRIBUTE_KEYS,
    NO_REQUIRED_KEYS,
  );
  if (!attributes) return undefined;
  const result: {
    kind: 'polar';
    position: CustomGeometryPoint;
    radiusGuide?: string;
    minRadius?: CustomGeometryValue;
    maxRadius?: CustomGeometryValue;
    angleGuide?: string;
    minAngle?: CustomGeometryValue;
    maxAngle?: CustomGeometryValue;
  } = { kind: 'polar', position };
  if (Object.hasOwn(attributes, 'gdRefR')) {
    const value = parseCustomGeometryToken(attributes.gdRefR);
    if (value === undefined) return undefined;
    result.radiusGuide = value;
  }
  if (Object.hasOwn(attributes, 'minR')) {
    const value = parseCustomGeometryValue(attributes.minR, false);
    if (value === undefined) return undefined;
    result.minRadius = value;
  }
  if (Object.hasOwn(attributes, 'maxR')) {
    const value = parseCustomGeometryValue(attributes.maxR, false);
    if (value === undefined) return undefined;
    result.maxRadius = value;
  }
  if (Object.hasOwn(attributes, 'gdRefAng')) {
    const value = parseCustomGeometryToken(attributes.gdRefAng);
    if (value === undefined) return undefined;
    result.angleGuide = value;
  }
  if (Object.hasOwn(attributes, 'minAng')) {
    const value = parseCustomGeometryValue(attributes.minAng, false);
    if (value === undefined) return undefined;
    result.minAngle = value;
  }
  if (Object.hasOwn(attributes, 'maxAng')) {
    const value = parseCustomGeometryValue(attributes.maxAng, false);
    if (value === undefined) return undefined;
    result.maxAngle = value;
  }
  return result;
}

function parseFormulaAttribute(value: string | undefined): CustomGeometryFormula | undefined {
  if (value === undefined) return undefined;
  const tokens = value.trim().split(/[ \t\r\n]+/);
  const operator = tokens.shift();
  if (!operator) return undefined;
  const arity = FORMULA_ARITIES.get(operator);
  if (arity === undefined || tokens.length !== arity) return undefined;
  const operands: CustomGeometryValue[] = [];
  for (const token of tokens) {
    const operand = parseCustomGeometryValue(token, false);
    if (operand === undefined) return undefined;
    operands.push(operand);
  }
  switch (operator) {
    case 'val':
    case 'abs':
    case 'sqrt':
      return { operator, operands: [operands[0]!] };
    case 'at2':
    case 'cos':
    case 'max':
    case 'min':
    case 'sin':
    case 'tan':
      return { operator, operands: [operands[0]!, operands[1]!] };
    case '*/':
    case '+-':
    case '+/':
    case '?:':
    case 'cat2':
    case 'mod':
    case 'pin':
    case 'sat2':
      return { operator, operands: [operands[0]!, operands[1]!, operands[2]!] };
    default:
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
      const widthRadius = parseCustomGeometryValue(attributes.wR, true);
      const heightRadius = parseCustomGeometryValue(attributes.hR, true);
      const startAngle = parseCustomGeometryValue(attributes.stAng, false);
      const sweepAngle = parseCustomGeometryValue(attributes.swAng, false);
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

function parsePointElement(
  element: XmlElement,
  expectedName = 'pt',
): CustomGeometryPoint | undefined {
  if (
    element.localName !== expectedName
    || namespaceUri(element) !== DRAWING_NAMESPACE
    || directChildren(element).length !== 0
    || hasNonWhitespaceText(element)
  ) return undefined;
  const attributes = readXmlAttributes(element, POINT_KEYS, POINT_KEYS);
  if (!attributes) return undefined;
  const x = parseCustomGeometryValue(attributes.x, false);
  const y = parseCustomGeometryValue(attributes.y, false);
  return x === undefined || y === undefined ? undefined : { x, y };
}

function parseCustomGeometryValue(
  value: string | undefined,
  positive: boolean,
): CustomGeometryValue | undefined {
  if (value === undefined) return undefined;
  if (INTEGER_PATTERN.test(value)) return parseInteger(value, positive);
  return parseCustomGeometryToken(value);
}

function parseCustomGeometryToken(value: string | undefined): string | undefined {
  if (value === undefined) return undefined;
  try {
    return normalizeCustomGeometryToken(value, 'Custom geometry token');
  } catch {
    return undefined;
  }
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

function guideListsEqual(
  left: readonly Readonly<CustomGeometryGuide>[] | undefined,
  right: readonly Readonly<CustomGeometryGuide>[] | undefined,
): boolean {
  if (left === undefined || right === undefined) return left === right;
  return left.length === right.length && left.every((guide, index) => {
    const other = right[index];
    return other !== undefined
      && guide.name === other.name
      && formulasEqual(guide.formula, other.formula);
  });
}

function formulasEqual(
  left: Readonly<CustomGeometryFormula>,
  right: Readonly<CustomGeometryFormula>,
): boolean {
  return left.operator === right.operator
    && left.operands.length === right.operands.length
    && left.operands.every((operand, index) => operand === right.operands[index]);
}

function handleListsEqual(
  left: readonly Readonly<CustomGeometryHandle>[] | undefined,
  right: readonly Readonly<CustomGeometryHandle>[] | undefined,
): boolean {
  if (left === undefined || right === undefined) return left === right;
  return left.length === right.length && left.every((handle, index) =>
    handlesEqual(handle, right[index]));
}

function handlesEqual(
  left: Readonly<CustomGeometryHandle>,
  right: Readonly<CustomGeometryHandle> | undefined,
): boolean {
  if (!right || left.kind !== right.kind || !pointsEqual(left.position, right.position)) {
    return false;
  }
  if (left.kind === 'xy') {
    if (right.kind !== 'xy') return false;
    return optionalPropertyEqual(left, right, 'xGuide')
      && optionalPropertyEqual(left, right, 'minX')
      && optionalPropertyEqual(left, right, 'maxX')
      && optionalPropertyEqual(left, right, 'yGuide')
      && optionalPropertyEqual(left, right, 'minY')
      && optionalPropertyEqual(left, right, 'maxY');
  }
  if (right.kind !== 'polar') return false;
  return optionalPropertyEqual(left, right, 'radiusGuide')
    && optionalPropertyEqual(left, right, 'minRadius')
    && optionalPropertyEqual(left, right, 'maxRadius')
    && optionalPropertyEqual(left, right, 'angleGuide')
    && optionalPropertyEqual(left, right, 'minAngle')
    && optionalPropertyEqual(left, right, 'maxAngle');
}

function connectionSiteListsEqual(
  left: readonly Readonly<CustomGeometryConnectionSite>[] | undefined,
  right: readonly Readonly<CustomGeometryConnectionSite>[] | undefined,
): boolean {
  if (left === undefined || right === undefined) return left === right;
  return left.length === right.length && left.every((site, index) => {
    const other = right[index];
    return other !== undefined
      && site.angle === other.angle
      && pointsEqual(site.position, other.position);
  });
}

function textRectanglesEqual(
  left: Readonly<CustomGeometryTextRectangle> | undefined,
  right: Readonly<CustomGeometryTextRectangle> | undefined,
): boolean {
  if (left === undefined || right === undefined) return left === right;
  return left.left === right.left
    && left.top === right.top
    && left.right === right.right
    && left.bottom === right.bottom;
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

function containsInvalidXmlCharacter(value: string): boolean {
  for (const character of value) {
    const codePoint = character.codePointAt(0)!;
    if (
      codePoint !== 0x09
      && codePoint !== 0x0a
      && codePoint !== 0x0d
      && (codePoint < 0x20
        || codePoint > 0x10ffff
        || (codePoint >= 0xd800 && codePoint <= 0xdfff)
        || codePoint === 0xfffe
        || codePoint === 0xffff)
    ) return true;
  }
  return false;
}
