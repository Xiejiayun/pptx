import { escapeXmlAttribute, type XmlElement } from '@pptx/lossless-xml';
import type { RichTextColor } from './text.js';

const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const EMU_PER_POINT = 12_700;
const ANGLE_SCALE = 60_000;
const OPACITY_SCALE = 100_000;
const MAX_BLUR_POINTS = 100;
const MAX_DISTANCE_POINTS = 200;
const MAX_ANGLE_DEGREES = 360;
const DEFAULT_COLOR: RichTextColor = { kind: 'srgb', value: '000000' };
const DEFAULT_OPACITY = 0.75;
const DEFAULT_BLUR = 8;
const DEFAULT_ANGLE = 270;
const DEFAULT_DISTANCE = 4;
const SCHEME_COLORS = new Set([
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'bg1',
  'bg2',
  'dk1',
  'dk2',
  'folHlink',
  'hlink',
  'lt1',
  'lt2',
  'phClr',
  'tx1',
  'tx2',
]);
const OUTER_ATTRIBUTE_NAMES = new Set([
  'sx',
  'sy',
  'kx',
  'ky',
  'algn',
  'rotWithShape',
  'blurRad',
  'dist',
  'dir',
]);
const INNER_ATTRIBUTE_NAMES = new Set(['blurRad', 'dist', 'dir']);
const OUTER_NEUTRAL_ATTRIBUTES = Object.freeze({
  sx: '100000',
  sy: '100000',
  kx: '0',
  ky: '0',
  algn: 'bl',
});

interface NormalizedShadowBase {
  readonly color: RichTextColor;
  readonly opacity: number;
  readonly blur: number;
  readonly angle: number;
  readonly distance: number;
}

export type NormalizedShapeShadow = Readonly<
  | (NormalizedShadowBase & {
      readonly kind: 'outer';
      readonly rotateWithShape: boolean;
    })
  | (NormalizedShadowBase & {
      readonly kind: 'inner';
    })
>;

export function normalizeShapeShadow(
  value: unknown,
  context: string,
): NormalizedShapeShadow {
  const candidate = readDataObject(value, context, [
    'kind',
    'color',
    'opacity',
    'blur',
    'angle',
    'distance',
    'rotateWithShape',
  ]);
  if (candidate.kind !== 'outer' && candidate.kind !== 'inner') {
    throw new TypeError(`${context} kind must be outer or inner`);
  }
  if (candidate.kind === 'inner' && candidate.rotateWithShape !== undefined) {
    throw new TypeError(`${context} inner shadow does not support rotateWithShape`);
  }
  if (
    candidate.kind === 'outer'
    && candidate.rotateWithShape !== undefined
    && typeof candidate.rotateWithShape !== 'boolean'
  ) {
    throw new TypeError(`${context} rotateWithShape must be a boolean`);
  }

  const color = candidate.color === undefined
    ? normalizeColor(DEFAULT_COLOR, `${context} color`)
    : normalizeColor(candidate.color, `${context} color`);
  const opacity = candidate.opacity === undefined
    ? DEFAULT_OPACITY
    : normalizeNumber(candidate.opacity, 0, 1, OPACITY_SCALE, `${context} opacity`);
  const blur = candidate.blur === undefined
    ? DEFAULT_BLUR
    : normalizeNumber(
      candidate.blur,
      0,
      MAX_BLUR_POINTS,
      EMU_PER_POINT,
      `${context} blur`,
    );
  const angle = candidate.angle === undefined
    ? DEFAULT_ANGLE
    : normalizeNumber(
      candidate.angle,
      0,
      MAX_ANGLE_DEGREES,
      ANGLE_SCALE,
      `${context} angle`,
      true,
    );
  const distance = candidate.distance === undefined
    ? DEFAULT_DISTANCE
    : normalizeNumber(
      candidate.distance,
      0,
      MAX_DISTANCE_POINTS,
      EMU_PER_POINT,
      `${context} distance`,
    );

  return freezeShadow(candidate.kind === 'outer'
    ? {
        kind: 'outer',
        color,
        opacity,
        blur,
        angle,
        distance,
        rotateWithShape: (candidate.rotateWithShape as boolean | undefined) ?? false,
      }
    : {
        kind: 'inner',
        color,
        opacity,
        blur,
        angle,
        distance,
      });
}

export function readSimpleShadow(
  element: XmlElement,
  prefix: string,
): NormalizedShapeShadow | undefined {
  const kind = element.localName === 'outerShdw'
    ? 'outer'
    : element.localName === 'innerShdw'
      ? 'inner'
      : undefined;
  if (
    !kind
    || element.name !== `${prefix}${element.localName}`
    || namespaceUri(element) !== DRAWING_NAMESPACE
  ) return undefined;

  const attributes = readAttributes(
    element,
    kind === 'outer' ? OUTER_ATTRIBUTE_NAMES : INNER_ATTRIBUTE_NAMES,
  );
  if (!attributes) return undefined;
  if (kind === 'outer' && !hasCompatibleOuterAttributes(attributes)) {
    return undefined;
  }

  const blurEmus = readInteger(attributes.get('blurRad'), MAX_BLUR_POINTS * EMU_PER_POINT);
  const distanceEmus = readInteger(
    attributes.get('dist'),
    MAX_DISTANCE_POINTS * EMU_PER_POINT,
  );
  const direction = readInteger(
    attributes.get('dir'),
    MAX_ANGLE_DEGREES * ANGLE_SCALE - 1,
  );
  if (blurEmus === undefined || distanceEmus === undefined || direction === undefined) {
    return undefined;
  }

  let rotateWithShape = false;
  if (kind === 'outer') {
    const raw = attributes.get('rotWithShape');
    if (raw === undefined) rotateWithShape = true;
    else if (raw === '1' || raw === 'true') rotateWithShape = true;
    else if (raw === '0' || raw === 'false') rotateWithShape = false;
    else return undefined;
  }

  const color = readShadowColor(element, prefix);
  if (!color) return undefined;
  return freezeShadow(kind === 'outer'
    ? {
        kind: 'outer',
        color: color.color,
        opacity: color.opacity,
        blur: blurEmus / EMU_PER_POINT,
        angle: direction / ANGLE_SCALE,
        distance: distanceEmus / EMU_PER_POINT,
        rotateWithShape,
      }
    : {
        kind: 'inner',
        color: color.color,
        opacity: color.opacity,
        blur: blurEmus / EMU_PER_POINT,
        angle: direction / ANGLE_SCALE,
        distance: distanceEmus / EMU_PER_POINT,
      });
}

export function renderSimpleShadow(
  shadow: NormalizedShapeShadow,
  prefix: string,
): string {
  const colorTag = shadow.color.kind === 'srgb' ? 'srgbClr' : 'schemeClr';
  const color = `<${prefix}${colorTag} val="${escapeXmlAttribute(shadow.color.value)}">` +
    `<${prefix}alpha val="${Math.round(shadow.opacity * OPACITY_SCALE)}"/>` +
    `</${prefix}${colorTag}>`;
  const geometry = `blurRad="${Math.round(shadow.blur * EMU_PER_POINT)}" ` +
    `dist="${Math.round(shadow.distance * EMU_PER_POINT)}" ` +
    `dir="${Math.round(shadow.angle * ANGLE_SCALE)}"`;
  if (shadow.kind === 'inner') {
    return `<${prefix}innerShdw ${geometry}>${color}</${prefix}innerShdw>`;
  }
  return `<${prefix}outerShdw sx="100000" sy="100000" kx="0" ky="0" ` +
    `algn="bl" rotWithShape="${shadow.rotateWithShape ? 1 : 0}" ${geometry}>` +
    `${color}</${prefix}outerShdw>`;
}

export function shapeShadowsEqual(
  left: NormalizedShapeShadow | undefined,
  right: NormalizedShapeShadow | undefined,
): boolean {
  if (!left || !right) return left === right;
  return left.kind === right.kind
    && left.color.kind === right.color.kind
    && left.color.value === right.color.value
    && left.opacity === right.opacity
    && left.blur === right.blur
    && left.angle === right.angle
    && left.distance === right.distance
    && (left.kind !== 'outer'
      || right.kind !== 'outer'
      || left.rotateWithShape === right.rotateWithShape);
}

function readShadowColor(
  shadow: XmlElement,
  prefix: string,
): { readonly color: RichTextColor; readonly opacity: number } | undefined {
  const children = directChildren(shadow);
  if (children.length !== 1) return undefined;
  const element = children[0]!;
  if (
    element.name !== `${prefix}${element.localName}`
    || namespaceUri(element) !== DRAWING_NAMESPACE
    || (element.localName !== 'srgbClr' && element.localName !== 'schemeClr')
  ) return undefined;
  const attributes = readAttributes(element, new Set(['val']));
  if (!attributes || attributes.size !== 1) return undefined;
  const value = attributes.get('val');
  let color: RichTextColor;
  if (element.localName === 'srgbClr') {
    if (!value || !/^[\da-f]{6}$/i.test(value)) return undefined;
    color = { kind: 'srgb', value: value.toUpperCase() };
  } else {
    if (!value || !SCHEME_COLORS.has(value)) return undefined;
    color = { kind: 'scheme', value };
  }

  const transforms = directChildren(element);
  if (transforms.length === 0) return { color, opacity: 1 };
  if (transforms.length !== 1) return undefined;
  const alpha = transforms[0]!;
  if (
    alpha.name !== `${prefix}alpha`
    || namespaceUri(alpha) !== DRAWING_NAMESPACE
    || directChildren(alpha).length !== 0
  ) return undefined;
  const alphaAttributes = readAttributes(alpha, new Set(['val']));
  if (!alphaAttributes || alphaAttributes.size !== 1) return undefined;
  const opacity = readInteger(alphaAttributes.get('val'), OPACITY_SCALE, false);
  return opacity === undefined ? undefined : { color, opacity: opacity / OPACITY_SCALE };
}

function readAttributes(
  element: XmlElement,
  allowed: ReadonlySet<string>,
): Map<string, string> | undefined {
  const result = new Map<string, string>();
  for (const attribute of element.attributes) {
    if (attribute.name === 'xmlns' || attribute.name.startsWith('xmlns:')) continue;
    if (!allowed.has(attribute.name) || result.has(attribute.name)) return undefined;
    result.set(attribute.name, attribute.value);
  }
  return result;
}

function hasCompatibleOuterAttributes(attributes: ReadonlyMap<string, string>): boolean {
  const names = Object.keys(OUTER_NEUTRAL_ATTRIBUTES) as Array<
    keyof typeof OUTER_NEUTRAL_ATTRIBUTES
  >;
  const present = names.filter((name) => attributes.has(name));
  if (present.length === 0) return true;
  return present.length === names.length
    && names.every((name) => attributes.get(name) === OUTER_NEUTRAL_ATTRIBUTES[name]);
}

function readInteger(
  raw: string | undefined,
  maximum: number,
  defaultZero = true,
): number | undefined {
  if (raw === undefined) return defaultZero ? 0 : undefined;
  if (!/^\d+$/.test(raw)) return undefined;
  const value = Number(raw);
  return Number.isSafeInteger(value) && value >= 0 && value <= maximum
    ? value
    : undefined;
}

function normalizeColor(value: unknown, context: string): RichTextColor {
  const candidate = readDataObject(value, context, ['kind', 'value']);
  if (candidate.kind === 'srgb') {
    if (typeof candidate.value !== 'string' || !/^#?[\da-f]{6}$/i.test(candidate.value)) {
      throw new TypeError(`${context} sRGB value must contain six hex digits`);
    }
    return Object.freeze({
      kind: 'srgb',
      value: candidate.value.replace(/^#/, '').toUpperCase(),
    });
  }
  if (candidate.kind === 'scheme') {
    if (typeof candidate.value !== 'string' || !SCHEME_COLORS.has(candidate.value)) {
      throw new TypeError(`${context} scheme value is unsupported`);
    }
    return Object.freeze({ kind: 'scheme', value: candidate.value });
  }
  throw new TypeError(`${context} kind must be srgb or scheme`);
}

function normalizeNumber(
  value: unknown,
  minimum: number,
  maximum: number,
  scale: number,
  context: string,
  exclusiveMaximum = false,
): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  if (value < minimum || (exclusiveMaximum ? value >= maximum : value > maximum)) {
    const upperBound = exclusiveMaximum ? `< ${maximum}` : `<= ${maximum}`;
    throw new RangeError(`${context} must be >= ${minimum} and ${upperBound}`);
  }
  return Math.round(value * scale) / scale;
}

function readDataObject(
  value: unknown,
  context: string,
  supported: readonly string[],
): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${context} must be an object`);
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError(`${context} must be an ordinary object`);
  }
  const allowed = new Set(supported);
  const result: Record<string, unknown> = {};
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !allowed.has(key)) {
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

function freezeShadow(
  shadow: NormalizedShapeShadow,
): NormalizedShapeShadow {
  const color = Object.freeze({ ...shadow.color }) as RichTextColor;
  return Object.freeze(shadow.kind === 'outer'
    ? { ...shadow, color }
    : { ...shadow, color });
}

function namespaceUri(element: XmlElement): string | undefined {
  const separator = element.name.indexOf(':');
  const prefix = separator < 0 ? '' : element.name.slice(0, separator);
  const declarationName = prefix === '' ? 'xmlns' : `xmlns:${prefix}`;
  for (let current: XmlElement | undefined = element; current; current = current.parent) {
    const declarations = current.attributes.filter(({ name }) => name === declarationName);
    if (declarations.length > 1) return undefined;
    if (declarations[0]) return declarations[0].value;
  }
  return undefined;
}

function directChildren(element: XmlElement): XmlElement[] {
  return element.children.filter(
    (child): child is XmlElement => child.type === 'element',
  );
}
