import {
  escapeXmlAttribute,
  LosslessXmlDocument,
  type XmlElement,
} from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
import type {
  TableCellBorder,
  TableCellBorders,
  TableCellBorderStyle,
} from './shapes.js';
import type { RichTextColor } from './text.js';

const EMU_PER_POINT = 12_700;
const MAX_LINE_WIDTH_EMU = 20_116_800;
const MAX_LINE_WIDTH_POINTS = MAX_LINE_WIDTH_EMU / EMU_PER_POINT;
const PUBLIC_SIDES = ['top', 'right', 'bottom', 'left'] as const;
const XML_SIDES = [
  ['left', 'lnL'],
  ['right', 'lnR'],
  ['top', 'lnT'],
  ['bottom', 'lnB'],
] as const;
const TAG_BY_SIDE = new Map(XML_SIDES);
const SCHEMA_ORDER = new Map<string, number>([
  ['lnL', 0],
  ['lnR', 1],
  ['lnT', 2],
  ['lnB', 3],
  ['lnTlToBr', 4],
  ['lnBlToTr', 5],
  ['noFill', 6],
  ['solidFill', 6],
  ['gradFill', 6],
  ['blipFill', 6],
  ['pattFill', 6],
  ['grpFill', 6],
  ['cell3D', 7],
  ['extLst', 8],
]);
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

type TableCellSide = typeof PUBLIC_SIDES[number];
type TableCellLineTag = typeof XML_SIDES[number][1];

export function normalizeTableCellBorders(
  value: unknown,
  context: string,
): TableCellBorders | undefined {
  if (value === undefined) return undefined;
  if (Array.isArray(value)) return normalizeTuple(value, context);
  if (!value || typeof value !== 'object') {
    throw new TypeError(`${context} must be a border, TRBL tuple, side object, or undefined`);
  }

  const candidate = value as Record<string, unknown>;
  if (Object.hasOwn(candidate, 'kind')) {
    const border = normalizeBorder(candidate, context);
    return {
      top: cloneBorder(border),
      right: cloneBorder(border),
      bottom: cloneBorder(border),
      left: cloneBorder(border),
    };
  }

  assertKeys(candidate, PUBLIC_SIDES, context);
  const result: Partial<Record<TableCellSide, TableCellBorder>> = {};
  for (const side of PUBLIC_SIDES) {
    if (!Object.hasOwn(candidate, side) || candidate[side] === undefined) continue;
    result[side] = normalizeBorder(candidate[side], `${context} ${side}`);
  }
  return Object.keys(result).length > 0 ? result : undefined;
}

export function readTableCellBorders(
  _xml: LosslessXmlDocument,
  cell: XmlElement,
): TableCellBorders | undefined {
  const directProperties = directChildren(cell).filter(({ localName }) => localName === 'tcPr');
  if (directProperties.length !== 1) return undefined;
  const properties = directProperties[0]!;
  const prefix = lexicalPrefix(properties.name);
  const result: Partial<Record<TableCellSide, TableCellBorder>> = {};

  for (const side of PUBLIC_SIDES) {
    const tag = TAG_BY_SIDE.get(side)!;
    const lines = directChildren(properties).filter(({ name }) => name === `${prefix}${tag}`);
    if (lines.length !== 1) continue;
    const border = readBorder(lines[0]!, prefix);
    if (border !== undefined) result[side] = border;
  }
  return Object.keys(result).length > 0 ? result : undefined;
}

export function replaceTableCellBorders(
  xml: LosslessXmlDocument,
  cell: XmlElement,
  borders: TableCellBorders | undefined,
  partUri: string,
): boolean {
  const directProperties = directChildren(cell).filter(({ localName }) => localName === 'tcPr');
  if (directProperties.length !== 1) {
    throw new ModelParseError(
      'Table cell must contain exactly one direct cell properties element',
      partUri,
    );
  }

  const propertiesElement = directProperties[0]!;
  const original = xml.original(propertiesElement);
  const properties = LosslessXmlDocument.parse(original);
  const root = requirePropertiesRoot(properties, partUri);
  const prefix = lexicalPrefix(root.name);
  const changes = new Set<TableCellSide>();

  for (const [side, tag] of XML_SIDES) {
    const lines = directChildren(root).filter(({ name }) => name === `${prefix}${tag}`);
    if (lines.length > 1) {
      throw new ModelParseError(`Table cell contains repeated direct ${tag} elements`, partUri);
    }
    const desired = borders?.[side];
    const current = lines[0] ? readBorder(lines[0], prefix) : undefined;
    const matches = lines.length === 0
      ? desired === undefined
      : desired !== undefined && bordersEqual(current, desired);
    if (!matches) changes.add(side);
  }
  if (changes.size === 0) return false;

  let updated = original;
  for (const [side, tag] of XML_SIDES) {
    if (!changes.has(side)) continue;
    updated = updateBorderSide(updated, tag, borders?.[side], partUri);
  }
  xml.replaceElement(propertiesElement, updated);
  return true;
}

function normalizeTuple(value: unknown[], context: string): TableCellBorders | undefined {
  if (value.length !== 4) {
    throw new TypeError(`${context} TRBL tuple must contain exactly four items`);
  }
  const allowed = new Set(['0', '1', '2', '3', 'length']);
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !allowed.has(key)) {
      throw new TypeError(`${context} TRBL tuple contains unsupported property ${String(key)}`);
    }
  }

  const result: Partial<Record<TableCellSide, TableCellBorder>> = {};
  for (let index = 0; index < PUBLIC_SIDES.length; index += 1) {
    if (!Object.hasOwn(value, index)) {
      throw new TypeError(`${context} TRBL tuple must not contain sparse items`);
    }
    const item = value[index];
    if (item !== undefined) {
      const side = PUBLIC_SIDES[index]!;
      result[side] = normalizeBorder(item, `${context} ${side}`);
    }
  }
  return Object.keys(result).length > 0 ? result : undefined;
}

function normalizeBorder(value: unknown, context: string): TableCellBorder {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${context} must be a border object`);
  }
  const candidate = value as Record<string, unknown>;
  if (candidate.kind === 'none') {
    assertKeys(candidate, ['kind'], context);
    return { kind: 'none' };
  }
  if (candidate.kind !== 'line') {
    throw new TypeError(`${context} kind must be none or line`);
  }
  assertKeys(candidate, ['kind', 'color', 'width', 'style'], context);
  if (candidate.color === undefined || candidate.width === undefined) {
    throw new TypeError(`${context} line must provide color and width`);
  }
  if (typeof candidate.width !== 'number' || !Number.isFinite(candidate.width)) {
    throw new TypeError(`${context} width must be finite`);
  }
  if (candidate.width < 0 || candidate.width > MAX_LINE_WIDTH_POINTS) {
    throw new RangeError(`${context} width must be between 0 and 1584 points`);
  }
  if (
    candidate.style !== undefined
    && candidate.style !== 'solid'
    && candidate.style !== 'dash'
  ) {
    throw new TypeError(`${context} style must be solid or dash`);
  }
  const widthEmu = Math.round(candidate.width * EMU_PER_POINT);
  if (widthEmu < 0 || widthEmu > MAX_LINE_WIDTH_EMU) {
    throw new RangeError(`${context} width must fit the OOXML line-width range`);
  }
  return {
    kind: 'line',
    color: normalizeColor(candidate.color, `${context} color`),
    width: widthEmu === 0 ? 0 : widthEmu / EMU_PER_POINT,
    ...(candidate.style !== undefined
      ? { style: candidate.style as TableCellBorderStyle }
      : {}),
  };
}

function normalizeColor(value: unknown, context: string): RichTextColor {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${context} must be an object`);
  }
  const candidate = value as Record<string, unknown>;
  assertKeys(candidate, ['kind', 'value'], context);
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

function readBorder(line: XmlElement, prefix: string): TableCellBorder | undefined {
  const attributes = nonNamespaceAttributes(line);
  const width = readStrictAttribute(attributes, 'w');
  if (width === undefined || !/^\d+$/.test(width)) return undefined;
  const widthEmu = Number(width);
  if (!Number.isSafeInteger(widthEmu) || widthEmu > MAX_LINE_WIDTH_EMU) return undefined;

  const allowedAttributes = new Set(['w', 'cap', 'cmpd', 'algn']);
  if (attributes.some(({ name }) => !allowedAttributes.has(name))) return undefined;
  if (
    readOptionalExactAttribute(attributes, 'cap', 'flat') === false
    || readOptionalExactAttribute(attributes, 'cmpd', 'sng') === false
    || readOptionalExactAttribute(attributes, 'algn', 'ctr') === false
  ) return undefined;

  const children = directChildren(line);
  if (children.some((child) => child.name !== `${prefix}${child.localName}`)) return undefined;
  if (children.length === 1 && children[0]?.localName === 'noFill') {
    const noFill = children[0];
    return noFill
      && nonNamespaceAttributes(noFill).length === 0
      && directChildren(noFill).length === 0
      ? { kind: 'none' }
      : undefined;
  }

  const allowedChildren = new Set(['solidFill', 'prstDash', 'round', 'headEnd', 'tailEnd']);
  if (children.some(({ localName }) => !allowedChildren.has(localName))) return undefined;
  const byName = new Map<string, XmlElement>();
  for (const child of children) {
    if (byName.has(child.localName)) return undefined;
    byName.set(child.localName, child);
  }
  const fill = byName.get('solidFill');
  if (!fill) return undefined;
  const color = readSolidColor(fill, prefix);
  if (!color) return undefined;

  let style: TableCellBorderStyle | undefined;
  const dash = byName.get('prstDash');
  if (dash) {
    const dashAttributes = nonNamespaceAttributes(dash);
    const value = readStrictAttribute(dashAttributes, 'val');
    if (
      dashAttributes.length !== 1
      || value === undefined
      || directChildren(dash).length > 0
    ) return undefined;
    if (value === 'solid') style = 'solid';
    else if (value === 'sysDash') style = 'dash';
    else return undefined;
  }

  const round = byName.get('round');
  if (round && (nonNamespaceAttributes(round).length > 0 || directChildren(round).length > 0)) {
    return undefined;
  }
  for (const tag of ['headEnd', 'tailEnd']) {
    const end = byName.get(tag);
    if (!end) continue;
    const endAttributes = nonNamespaceAttributes(end);
    if (
      endAttributes.length !== 3
      || readStrictAttribute(endAttributes, 'type') !== 'none'
      || readStrictAttribute(endAttributes, 'w') !== 'med'
      || readStrictAttribute(endAttributes, 'len') !== 'med'
      || directChildren(end).length > 0
    ) return undefined;
  }

  return {
    kind: 'line',
    color,
    width: widthEmu === 0 ? 0 : widthEmu / EMU_PER_POINT,
    ...(style !== undefined ? { style } : {}),
  };
}

function readSolidColor(fill: XmlElement, prefix: string): RichTextColor | undefined {
  if (nonNamespaceAttributes(fill).length > 0) return undefined;
  const children = directChildren(fill);
  if (children.length !== 1) return undefined;
  const colorElement = children[0]!;
  if (
    colorElement.name !== `${prefix}${colorElement.localName}`
    || (colorElement.localName !== 'srgbClr' && colorElement.localName !== 'schemeClr')
    || directChildren(colorElement).length > 0
  ) return undefined;
  const attributes = nonNamespaceAttributes(colorElement);
  const value = readStrictAttribute(attributes, 'val');
  if (attributes.length !== 1 || value === undefined) return undefined;
  if (colorElement.localName === 'srgbClr') {
    return /^[\da-f]{6}$/i.test(value)
      ? { kind: 'srgb', value: value.toUpperCase() }
      : undefined;
  }
  return SCHEME_COLORS.has(value) ? { kind: 'scheme', value } : undefined;
}

function updateBorderSide(
  template: string,
  tag: TableCellLineTag,
  border: TableCellBorder | undefined,
  partUri: string,
): string {
  const properties = LosslessXmlDocument.parse(template);
  const root = requirePropertiesRoot(properties, partUri);
  const prefix = lexicalPrefix(root.name);
  const lines = directChildren(root).filter(({ name }) => name === `${prefix}${tag}`);
  if (lines.length > 1) {
    throw new ModelParseError(`Table cell contains repeated direct ${tag} elements`, partUri);
  }
  const line = lines[0];
  if (border === undefined) {
    if (line) properties.removeElement(line);
    return properties.serialize();
  }

  const encoded = renderBorder(tag, border, prefix);
  if (line) {
    properties.replaceElement(line, encoded);
    return properties.serialize();
  }

  const rank = SCHEMA_ORDER.get(tag)!;
  const anchor = directChildren(root).find((child) => {
    if (child.name !== `${prefix}${child.localName}`) return false;
    const childRank = SCHEMA_ORDER.get(child.localName);
    return childRank !== undefined && childRank > rank;
  });
  if (anchor) properties.replace(anchor.start, anchor.start, encoded);
  else properties.appendChildXml(root, encoded);
  return properties.serialize();
}

function renderBorder(
  tag: TableCellLineTag,
  border: TableCellBorder,
  prefix: string,
): string {
  const width = border.kind === 'none' ? 0 : Math.round(border.width * EMU_PER_POINT);
  const open = `<${prefix}${tag} w="${width}" cap="flat" cmpd="sng" algn="ctr">`;
  if (border.kind === 'none') {
    return `${open}<${prefix}noFill/></${prefix}${tag}>`;
  }
  const colorTag = border.color.kind === 'srgb' ? 'srgbClr' : 'schemeClr';
  const color = escapeXmlAttribute(border.color.value);
  const dash = border.style === undefined
    ? ''
    : `<${prefix}prstDash val="${border.style === 'dash' ? 'sysDash' : 'solid'}"/>`;
  return `${open}<${prefix}solidFill><${prefix}${colorTag} val="${color}"/></${prefix}solidFill>${dash}<${prefix}round/><${prefix}headEnd type="none" w="med" len="med"/><${prefix}tailEnd type="none" w="med" len="med"/></${prefix}${tag}>`;
}

function bordersEqual(
  left: TableCellBorder | undefined,
  right: TableCellBorder,
): boolean {
  if (!left || left.kind !== right.kind) return false;
  if (left.kind === 'none' || right.kind === 'none') return true;
  return left.color.kind === right.color.kind
    && left.color.value === right.color.value
    && left.width === right.width
    && left.style === right.style;
}

function cloneBorder(border: TableCellBorder): TableCellBorder {
  return border.kind === 'none'
    ? { kind: 'none' }
    : {
        kind: 'line',
        color: { ...border.color },
        width: border.width,
        ...(border.style !== undefined ? { style: border.style } : {}),
      };
}

function readStrictAttribute(
  attributes: ReturnType<typeof nonNamespaceAttributes>,
  name: string,
): string | undefined {
  const matches = attributes.filter((attribute) => attribute.name === name);
  return matches.length === 1 ? matches[0]!.value : undefined;
}

function readOptionalExactAttribute(
  attributes: ReturnType<typeof nonNamespaceAttributes>,
  name: string,
  expected: string,
): boolean {
  const matches = attributes.filter((attribute) => attribute.name === name);
  return matches.length === 0 || (matches.length === 1 && matches[0]!.value === expected);
}

function assertKeys(value: object, supported: readonly string[], context: string): void {
  const allowed = new Set(supported);
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !allowed.has(key)) {
      throw new TypeError(`${context} contains unsupported property ${String(key)}`);
    }
  }
}

function requirePropertiesRoot(
  xml: LosslessXmlDocument,
  partUri: string,
): XmlElement {
  const root = xml.roots[0];
  if (!root || root.localName !== 'tcPr') {
    throw new ModelParseError('Invalid table cell properties template', partUri);
  }
  return root;
}

function nonNamespaceAttributes(element: XmlElement) {
  return element.attributes.filter(
    ({ name }) => name !== 'xmlns' && !name.startsWith('xmlns:'),
  );
}

function lexicalPrefix(name: string): string {
  const separator = name.indexOf(':');
  return separator < 0 ? '' : name.slice(0, separator + 1);
}

function directChildren(element: XmlElement): XmlElement[] {
  return element.children.filter(
    (child): child is XmlElement => child.type === 'element',
  );
}
