import type { XmlElement } from '@pptx/lossless-xml';
import type { RichTextColor } from './text.js';
import type { ShapeLineCap } from './preset-shape.js';
import {
  normalizeSimpleFill,
  readSimpleFillChoice,
  renderSimpleFill,
  SIMPLE_FILL_CHOICE_NAMES,
  simpleFillsEqual,
  type SimpleFill,
} from './simple-fill.internal.js';

const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const EMU_PER_POINT = 12_700;
const MAX_WIDTH_POINTS = 1_584;
const MAX_WIDTH_EMUS = MAX_WIDTH_POINTS * EMU_PER_POINT;

export type SimpleLineDash =
  | 'solid'
  | 'dash'
  | 'dashDot'
  | 'lgDash'
  | 'lgDashDot'
  | 'lgDashDotDot'
  | 'sysDash'
  | 'sysDot';

export type NormalizedSimpleLine =
  | { readonly kind: 'none' }
  | {
      readonly kind: 'line';
      readonly color: RichTextColor;
      readonly transparency?: number;
      readonly width: number;
      readonly dash: SimpleLineDash;
      readonly cap?: ShapeLineCap;
    };

export const SIMPLE_LINE_FILL_CHOICE_NAMES = SIMPLE_FILL_CHOICE_NAMES;
export const SIMPLE_LINE_DASH_CHOICE_NAMES = Object.freeze([
  'prstDash',
  'custDash',
] as const);

const DASHES = new Set<SimpleLineDash>([
  'solid',
  'dash',
  'dashDot',
  'lgDash',
  'lgDashDot',
  'lgDashDotDot',
  'sysDash',
  'sysDot',
]);
const FILL_CHOICES = new Set<string>(SIMPLE_LINE_FILL_CHOICE_NAMES);
const DASH_CHOICES = new Set<string>(SIMPLE_LINE_DASH_CHOICE_NAMES);
const LINE_CAP_TO_OOXML: Readonly<Record<ShapeLineCap, string>> = Object.freeze({
  flat: 'flat',
  round: 'rnd',
  square: 'sq',
});
const OOXML_TO_LINE_CAP: Readonly<Record<string, ShapeLineCap>> = Object.freeze({
  flat: 'flat',
  rnd: 'round',
  sq: 'square',
});

export function normalizeSimpleLine(
  value: unknown,
  context: string,
): NormalizedSimpleLine | undefined {
  if (value === undefined) return undefined;
  const candidate = readDataObject(
    value,
    context,
    ['kind', 'color', 'transparency', 'width', 'dash', 'cap'],
  );
  if (candidate.kind === 'none') {
    assertKeys(candidate, ['kind'], context);
    return { kind: 'none' };
  }
  if (candidate.kind !== 'line') {
    throw new TypeError(`${context} kind must be none or line`);
  }
  if (candidate.color === undefined) {
    throw new TypeError(`${context} line must provide color`);
  }

  const fill = normalizeSimpleFill({
    kind: 'solid',
    color: candidate.color,
    ...(candidate.transparency !== undefined
      ? { transparency: candidate.transparency }
      : {}),
  }, context);
  if (!fill || fill.kind !== 'solid') {
    throw new TypeError(`${context} line must provide a solid color`);
  }

  const width = candidate.width === undefined
    ? 1
    : normalizeWidth(candidate.width, `${context} width`);
  const dash = candidate.dash === undefined
    ? 'solid'
    : normalizeDash(candidate.dash, `${context} dash`);
  const cap = candidate.cap === undefined
    ? undefined
    : normalizeCap(candidate.cap, `${context} cap`);
  return {
    kind: 'line',
    color: fill.color,
    ...(fill.transparency !== undefined
      ? { transparency: fill.transparency }
      : {}),
    width,
    dash,
    ...(cap === undefined ? {} : { cap }),
  };
}

export function readSimpleLine(
  lineElement: XmlElement,
  prefix: string,
): NormalizedSimpleLine | undefined {
  if (
    lineElement.name !== `${prefix}ln`
    || namespaceUri(lineElement) !== DRAWING_NAMESPACE
  ) return undefined;

  let width = 1;
  let hasWidth = false;
  let cap: ShapeLineCap | undefined;
  const seenAttributes = new Set<string>();
  for (const attribute of nonNamespaceAttributes(lineElement)) {
    if (seenAttributes.has(attribute.name)) return undefined;
    seenAttributes.add(attribute.name);
    if (attribute.name === 'w') {
      if (!/^\d+$/.test(attribute.value)) return undefined;
      const emus = Number(attribute.value);
      if (!Number.isSafeInteger(emus) || emus > MAX_WIDTH_EMUS) return undefined;
      width = emus / EMU_PER_POINT;
      hasWidth = true;
      continue;
    }
    if (attribute.name === 'cap') {
      cap = OOXML_TO_LINE_CAP[attribute.value];
      if (cap === undefined) return undefined;
      continue;
    }
    if (attribute.name === 'cmpd' && attribute.value === 'sng') continue;
    if (attribute.name === 'algn' && attribute.value === 'ctr') continue;
    return undefined;
  }

  const children = directChildren(lineElement);
  const fillChoices = children.filter(({ localName }) => FILL_CHOICES.has(localName));
  if (fillChoices.length !== 1) return undefined;
  const fillChoice = fillChoices[0]!;
  if (!subtreeUsesDrawingNamespace(fillChoice)) return undefined;
  const fill = readSimpleFillChoice(fillChoice, prefix);
  if (!fill) return undefined;

  const dashChoices = children.filter(({ localName }) => DASH_CHOICES.has(localName));
  if (dashChoices.length > 1) return undefined;
  if (fill.kind === 'none') {
    return !hasWidth && cap === undefined && dashChoices.length === 0
      ? { kind: 'none' }
      : undefined;
  }

  let dash: SimpleLineDash = 'solid';
  const dashChoice = dashChoices[0];
  if (dashChoice) {
    if (
      dashChoice.localName !== 'prstDash'
      || dashChoice.name !== `${prefix}prstDash`
      || namespaceUri(dashChoice) !== DRAWING_NAMESPACE
      || directChildren(dashChoice).length !== 0
    ) return undefined;
    const attributes = nonNamespaceAttributes(dashChoice);
    if (attributes.length !== 1 || attributes[0]?.name !== 'val') return undefined;
    if (attributes[0].value === 'dot') dash = 'sysDot';
    else {
      if (!DASHES.has(attributes[0].value as SimpleLineDash)) return undefined;
      dash = attributes[0].value as SimpleLineDash;
    }
  }

  return {
    kind: 'line',
    color: { ...fill.color },
    ...(fill.transparency !== undefined
      ? { transparency: fill.transparency }
      : {}),
    width,
    dash,
    ...(cap === undefined ? {} : { cap }),
  };
}

export function renderSimpleLineAttributes(line: NormalizedSimpleLine): string {
  if (line.kind === 'none' || line.cap === undefined) return '';
  return ` cap="${LINE_CAP_TO_OOXML[line.cap]}"`;
}

export function renderSimpleLine(
  line: NormalizedSimpleLine,
  prefix: string,
): string {
  if (line.kind === 'none') return renderSimpleFill(line, prefix);
  const fill: SimpleFill = {
    kind: 'solid',
    color: line.color,
    ...(line.transparency !== undefined
      ? { transparency: line.transparency }
      : {}),
  };
  return renderSimpleFill(fill, prefix)
    + `<${prefix}prstDash val="${line.dash}"/>`;
}

export function simpleLinesEqual(
  left: NormalizedSimpleLine | undefined,
  right: NormalizedSimpleLine,
): boolean {
  if (!left || left.kind !== right.kind) return false;
  if (left.kind === 'none' || right.kind === 'none') return true;
  const leftFill: SimpleFill = {
    kind: 'solid',
    color: left.color,
    ...(left.transparency !== undefined
      ? { transparency: left.transparency }
      : {}),
  };
  const rightFill: SimpleFill = {
    kind: 'solid',
    color: right.color,
    ...(right.transparency !== undefined
      ? { transparency: right.transparency }
      : {}),
  };
  return simpleFillsEqual(leftFill, rightFill)
    && left.width === right.width
    && left.dash === right.dash
    && (left.cap ?? 'flat') === (right.cap ?? 'flat');
}

function normalizeWidth(value: unknown, context: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  if (value < 0 || value > MAX_WIDTH_POINTS) {
    throw new RangeError(`${context} must be between 0 and ${MAX_WIDTH_POINTS}`);
  }
  return Math.round(value * EMU_PER_POINT) / EMU_PER_POINT;
}

function normalizeDash(value: unknown, context: string): SimpleLineDash {
  if (typeof value !== 'string' || !DASHES.has(value as SimpleLineDash)) {
    throw new TypeError(`${context} is unsupported`);
  }
  return value as SimpleLineDash;
}

function normalizeCap(value: unknown, context: string): ShapeLineCap {
  if (typeof value !== 'string' || !Object.hasOwn(LINE_CAP_TO_OOXML, value)) {
    throw new TypeError(`${context} is unsupported`);
  }
  return value as ShapeLineCap;
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

function assertKeys(value: object, supported: readonly string[], context: string): void {
  const allowed = new Set(supported);
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !allowed.has(key)) {
      throw new TypeError(`${context} contains unsupported property ${String(key)}`);
    }
  }
}

function subtreeUsesDrawingNamespace(element: XmlElement): boolean {
  if (namespaceUri(element) !== DRAWING_NAMESPACE) return false;
  return directChildren(element).every(subtreeUsesDrawingNamespace);
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

function nonNamespaceAttributes(element: XmlElement) {
  return element.attributes.filter(
    ({ name }) => name !== 'xmlns' && !name.startsWith('xmlns:'),
  );
}

function directChildren(element: XmlElement): XmlElement[] {
  return element.children.filter(
    (child): child is XmlElement => child.type === 'element',
  );
}
