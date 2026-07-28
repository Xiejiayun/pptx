import {
  escapeXmlAttribute,
  LosslessXmlDocument,
  type XmlElement,
} from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
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
  if (Array.isArray(value)) {
    if (value.length !== 4) throw new RangeError(`${context} tuple must contain exactly four values`);
    return {
      top: normalizeSide(value[0], `${context} top`),
      right: normalizeSide(value[1], `${context} right`),
      bottom: normalizeSide(value[2], `${context} bottom`),
      left: normalizeSide(value[3], `${context} left`),
    };
  }
  if (!value || typeof value !== 'object') {
    throw new TypeError(`${context} must be a number, four-value tuple, or margin object`);
  }
  for (const key of Object.keys(value)) {
    if (!SIDE_NAMES.has(key as keyof TextBoxMargins)) {
      throw new TypeError(`${context} contains unsupported property ${key}`);
    }
  }
  const candidate = value as TextBoxMargins;
  const normalized: { top?: number; right?: number; bottom?: number; left?: number } = {};
  for (const [side] of SIDES) {
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
  const bodyProperties = requireBodyProperties(xml, shape, partUri);
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
  const bodyProperties = requireBodyProperties(xml, shape, partUri);
  let updated = xml.original(bodyProperties);
  for (const [side, attribute] of SIDES) {
    const value = margins?.[side];
    updated = updateAttribute(updated, attribute, value === undefined ? undefined : String(toRaw(value)));
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

function toRaw(value: number): number {
  return Math.round(value * EMU_PER_POINT);
}

function requireBodyProperties(
  xml: LosslessXmlDocument,
  shape: XmlElement,
  partUri: string,
): XmlElement {
  const textBody = directChildren(shape, 'txBody')[0];
  if (!textBody) throw new ModelParseError('Shape does not contain a direct text body', partUri);
  const bodyProperties = directChildren(textBody, 'bodyPr')[0];
  if (!bodyProperties) throw new ModelParseError('Shape does not contain direct text body properties', partUri);
  return bodyProperties;
}

function updateAttribute(template: string, name: string, value: string | undefined): string {
  const xml = LosslessXmlDocument.parse(template);
  const root = xml.roots[0];
  if (!root || root.localName !== 'bodyPr') throw new ModelParseError('Invalid text body properties template');
  const attribute = xml.attribute(root, name);
  if (value !== undefined) {
    if (attribute) xml.replaceAttribute(attribute, value);
    else {
      const insertionPoint = root.selfClosing
        ? xml.source.lastIndexOf('/', root.startTagEnd - 1)
        : root.startTagEnd - 1;
      xml.replace(insertionPoint, insertionPoint, ` ${name}="${escapeXmlAttribute(value)}"`);
    }
  } else if (attribute) {
    let start = attribute.start;
    while (start > root.start && /[\t ]/.test(xml.source[start - 1] ?? '')) start -= 1;
    xml.replace(start, attribute.end, '');
  }
  return xml.serialize();
}

function directChildren(element: XmlElement, localName: string): XmlElement[] {
  return element.children.filter(
    (child): child is XmlElement => child.type === 'element' && child.localName === localName,
  );
}
