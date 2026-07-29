import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import {
  requireTextBodyProperties,
  updateTextBodyAttribute,
} from './text-body-properties.internal.js';
import type { TextBoxMarginInput, TextBoxMargins } from './text.js';

const EMU_PER_POINT = 12_700;
const MIN_INT32 = -2_147_483_648;
const MAX_INT32 = 2_147_483_647;
const SIDES = [
  ['left', 'lIns'],
  ['top', 'tIns'],
  ['right', 'rIns'],
  ['bottom', 'bIns'],
] as const;
const SIDE_NAMES = new Set(SIDES.map(([side]) => side));

export function normalizeTextBoxMargins(
  value: TextBoxMarginInput | undefined,
  context: string,
): TextBoxMargins | undefined {
  if (value === undefined) return undefined;
  if (typeof value === 'number') {
    const margin = normalizeSide(value, context);
    return { top: margin, right: margin, bottom: margin, left: margin };
  }
  if (Array.isArray(value)) return normalizeTuple(value, context);
  const candidate = readMarginObject(value, context);
  const normalized: { top?: number; right?: number; bottom?: number; left?: number } = {};
  for (const [side] of SIDES) {
    if (!Object.hasOwn(candidate, side)) continue;
    const margin = candidate[side];
    if (margin !== undefined) normalized[side] = normalizeSide(margin, `${context} ${side}`);
  }
  return normalized;
}

export function renderTextBoxMarginAttributes(margins: TextBoxMargins | undefined): string {
  if (!margins) return '';
  return SIDES.map(([side, attribute]) => {
    const value = margins[side];
    return value === undefined ? '' : ` ${attribute}="${toRaw(value)}"`;
  }).join('');
}

export function readTextBoxMargins(
  xml: LosslessXmlDocument,
  shape: XmlElement,
  partUri: string,
): TextBoxMargins | undefined {
  const bodyProperties = requireTextBodyProperties(xml, shape, partUri);
  const margins: { top?: number; right?: number; bottom?: number; left?: number } = {};
  for (const [side, attributeName] of SIDES) {
    const attribute = xml.attribute(bodyProperties, attributeName);
    if (!attribute || !/^-?\d+$/.test(attribute.value)) continue;
    const raw = Number(attribute.value);
    if (!Number.isSafeInteger(raw) || raw < MIN_INT32 || raw > MAX_INT32) continue;
    margins[side] = raw / EMU_PER_POINT;
  }
  return Object.keys(margins).length > 0 ? margins : undefined;
}

export function replaceTextBoxMargins(
  xml: LosslessXmlDocument,
  shape: XmlElement,
  margins: TextBoxMargins | undefined,
  partUri: string,
): void {
  const bodyProperties = requireTextBodyProperties(xml, shape, partUri);
  let updated = xml.original(bodyProperties);
  for (const [side, attribute] of SIDES) {
    const value = margins?.[side];
    updated = updateTextBodyAttribute(
      updated,
      attribute,
      value === undefined ? undefined : String(toRaw(value)),
    );
  }
  xml.replaceElement(bodyProperties, updated);
}

function normalizeSide(value: unknown, context: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  const raw = Math.round(value * EMU_PER_POINT);
  if (raw < MIN_INT32 || raw > MAX_INT32) {
    throw new RangeError(`${context} must fit the OOXML signed Int32 coordinate range`);
  }
  return raw / EMU_PER_POINT;
}

function normalizeTuple(value: unknown[], context: string): TextBoxMargins {
  if (Object.getPrototypeOf(value) !== Array.prototype) {
    throw new TypeError(`${context} tuple must be an ordinary array`);
  }
  const length = Object.getOwnPropertyDescriptor(value, 'length');
  if (!length || !Object.hasOwn(length, 'value') || length.value !== 4) {
    throw new RangeError(`${context} tuple must contain exactly four values`);
  }
  const allowed = new Set(['0', '1', '2', '3', 'length']);
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !allowed.has(key)) {
      throw new TypeError(`${context} tuple contains unsupported property ${String(key)}`);
    }
  }
  const values = Array.from({ length: 4 }, (_, index) => {
    const descriptor = Object.getOwnPropertyDescriptor(value, String(index));
    if (!descriptor) throw new TypeError(`${context} tuple must not contain sparse values`);
    if (!Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`${context} tuple must contain only data values`);
    }
    return descriptor.value;
  });
  return {
    top: normalizeSide(values[0], `${context} top`),
    right: normalizeSide(values[1], `${context} right`),
    bottom: normalizeSide(values[2], `${context} bottom`),
    left: normalizeSide(values[3], `${context} left`),
  };
}

function readMarginObject(value: unknown, context: string): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${context} must be a number, four-value tuple, or margin object`);
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError(`${context} must be an ordinary object`);
  }
  const result: Record<string, unknown> = {};
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !SIDE_NAMES.has(key as keyof TextBoxMargins)) {
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

function toRaw(value: number): number {
  return Math.round(value * EMU_PER_POINT);
}
