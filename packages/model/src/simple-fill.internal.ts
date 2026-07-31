import { escapeXmlAttribute, type XmlElement } from '@pptx/lossless-xml';
import type { RichTextColor } from './text.js';

export type SimpleFill =
  | { readonly kind: 'none' }
  | {
      readonly kind: 'solid';
      readonly color: RichTextColor;
      readonly transparency?: number;
    };

export const SIMPLE_FILL_CHOICE_NAMES = Object.freeze([
  'noFill',
  'solidFill',
  'gradFill',
  'blipFill',
  'pattFill',
  'grpFill',
] as const);

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

export function normalizeSimpleFill(
  value: unknown,
  context: string,
): SimpleFill | undefined {
  if (value === undefined) return undefined;
  const candidate = readDataObject(
    value,
    context,
    ['kind', 'color', 'transparency'],
  );
  if (candidate.kind === 'none') {
    assertKeys(candidate, ['kind'], context);
    return { kind: 'none' };
  }
  if (candidate.kind !== 'solid') {
    throw new TypeError(`${context} kind must be none or solid`);
  }
  if (candidate.color === undefined) {
    throw new TypeError(`${context} solid fill must provide color`);
  }
  const color = normalizeColor(candidate.color, `${context} color`);
  const transparency = candidate.transparency === undefined
    ? undefined
    : normalizeTransparency(candidate.transparency, `${context} transparency`);
  return {
    kind: 'solid',
    color,
    ...(transparency !== undefined ? { transparency } : {}),
  };
}

export function readSimpleFillChoice(
  choice: XmlElement,
  prefix: string,
): SimpleFill | undefined {
  if (choice.name !== `${prefix}${choice.localName}`) return undefined;
  if (choice.localName === 'noFill') {
    return nonNamespaceAttributes(choice).length === 0 && directChildren(choice).length === 0
      ? { kind: 'none' }
      : undefined;
  }
  if (choice.localName !== 'solidFill' || nonNamespaceAttributes(choice).length !== 0) {
    return undefined;
  }

  const colorChildren = directChildren(choice);
  if (colorChildren.length !== 1) return undefined;
  const colorElement = colorChildren[0]!;
  if (
    colorElement.name !== `${prefix}${colorElement.localName}`
    || (colorElement.localName !== 'srgbClr' && colorElement.localName !== 'schemeClr')
  ) return undefined;

  const colorAttributes = nonNamespaceAttributes(colorElement);
  if (colorAttributes.length !== 1 || colorAttributes[0]?.name !== 'val') return undefined;
  const rawColor = colorAttributes[0].value;
  let color: RichTextColor;
  if (colorElement.localName === 'srgbClr') {
    if (!/^[\da-f]{6}$/i.test(rawColor)) return undefined;
    color = { kind: 'srgb', value: rawColor.toUpperCase() };
  } else {
    if (!SCHEME_COLORS.has(rawColor)) return undefined;
    color = { kind: 'scheme', value: rawColor };
  }

  const transforms = directChildren(colorElement);
  if (transforms.length === 0) return { kind: 'solid', color };
  if (transforms.length !== 1) return undefined;
  const alpha = transforms[0]!;
  if (alpha.name !== `${prefix}alpha` || directChildren(alpha).length > 0) return undefined;
  const alphaAttributes = nonNamespaceAttributes(alpha);
  if (alphaAttributes.length !== 1 || alphaAttributes[0]?.name !== 'val') return undefined;
  const rawAlpha = alphaAttributes[0].value;
  if (!/^\d+$/.test(rawAlpha)) return undefined;
  const alphaValue = Number(rawAlpha);
  if (!Number.isSafeInteger(alphaValue) || alphaValue < 0 || alphaValue > 100_000) {
    return undefined;
  }
  return {
    kind: 'solid',
    color,
    transparency: 100 - alphaValue / 1_000,
  };
}

export function renderSimpleFill(fill: SimpleFill, prefix: string): string {
  if (fill.kind === 'none') return `<${prefix}noFill/>`;
  const tag = fill.color.kind === 'srgb' ? 'srgbClr' : 'schemeClr';
  const value = escapeXmlAttribute(fill.color.value);
  const color = fill.transparency === undefined
    ? `<${prefix}${tag} val="${value}"/>`
    : `<${prefix}${tag} val="${value}"><${prefix}alpha val="${Math.round(
        (100 - fill.transparency) * 1_000,
      )}"/></${prefix}${tag}>`;
  return `<${prefix}solidFill>${color}</${prefix}solidFill>`;
}

export function simpleFillsEqual(
  left: SimpleFill | undefined,
  right: SimpleFill,
): boolean {
  if (!left || left.kind !== right.kind) return false;
  if (left.kind === 'none' || right.kind === 'none') return true;
  return left.color.kind === right.color.kind
    && left.color.value === right.color.value
    && left.transparency === right.transparency;
}

function normalizeColor(value: unknown, context: string): RichTextColor {
  const candidate = readDataObject(value, context, ['kind', 'value']);
  if (candidate.kind === 'srgb') {
    if (typeof candidate.value !== 'string' || !/^#?[\da-f]{6}$/i.test(candidate.value)) {
      throw new TypeError(`${context} sRGB value must contain six hex digits`);
    }
    return { kind: 'srgb', value: candidate.value.replace(/^#/, '').toUpperCase() };
  }
  if (candidate.kind === 'scheme') {
    if (typeof candidate.value !== 'string' || !SCHEME_COLORS.has(candidate.value)) {
      throw new TypeError(`${context} scheme value is unsupported`);
    }
    return { kind: 'scheme', value: candidate.value };
  }
  throw new TypeError(`${context} kind must be srgb or scheme`);
}

function normalizeTransparency(value: unknown, context: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  if (value < 0 || value > 100) {
    throw new RangeError(`${context} must be between 0 and 100`);
  }
  return (100_000 - Math.round((100 - value) * 1_000)) / 1_000;
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
