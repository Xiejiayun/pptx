import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
import type { TextBoxMargins } from './text.js';

const EMU_PER_POINT = 12_700;
const MIN_INT32 = -2_147_483_648;
const MAX_INT32 = 2_147_483_647;
const SIDES = [
  ['left', 'marL'],
  ['right', 'marR'],
  ['top', 'marT'],
  ['bottom', 'marB'],
] as const;
const DEFAULT_MARGINS: Required<TextBoxMargins> = {
  top: 3.6,
  right: 7.2,
  bottom: 3.6,
  left: 7.2,
};

export function renderTableCellMarginAttributes(
  margins: TextBoxMargins | undefined,
): string {
  return SIDES.map(([side, attribute]) => {
    const points = margins?.[side] ?? DEFAULT_MARGINS[side];
    return ` ${attribute}="${Math.round(points * EMU_PER_POINT)}"`;
  }).join('');
}

export function readTableCellMargins(
  _xml: LosslessXmlDocument,
  cell: XmlElement,
): TextBoxMargins | undefined {
  const directProperties = directChildren(cell, 'tcPr');
  if (directProperties.length !== 1) return undefined;

  const margins: { top?: number; right?: number; bottom?: number; left?: number } = {};
  for (const [side, attributeName] of SIDES) {
    const attributes = directProperties[0]!.attributes.filter(
      ({ name }) => name === attributeName,
    );
    if (attributes.length !== 1) continue;
    const raw = parseRaw(attributes[0]!.value);
    if (raw !== undefined) margins[side] = raw / EMU_PER_POINT;
  }
  return Object.keys(margins).length > 0 ? margins : undefined;
}

export function replaceTableCellMargins(
  xml: LosslessXmlDocument,
  cell: XmlElement,
  margins: TextBoxMargins | undefined,
  partUri: string,
): boolean {
  const directProperties = directChildren(cell, 'tcPr');
  if (directProperties.length !== 1) {
    throw new ModelParseError(
      'Table cell must contain exactly one direct cell properties element',
      partUri,
    );
  }

  const propertiesElement = directProperties[0]!;
  const original = xml.original(propertiesElement);
  const properties = LosslessXmlDocument.parse(original);
  const root = properties.roots[0];
  if (!root || root.localName !== 'tcPr') {
    throw new ModelParseError('Invalid table cell properties template', partUri);
  }

  const updates: Array<readonly [string, string | undefined]> = [];
  for (const [side, attributeName] of SIDES) {
    const attributes = root.attributes.filter(({ name }) => name === attributeName);
    if (attributes.length > 1) {
      throw new ModelParseError(
        `Table cell contains repeated direct ${attributeName} attributes`,
        partUri,
      );
    }
    const value = margins?.[side];
    const token = value === undefined ? undefined : String(Math.round(value * EMU_PER_POINT));
    const current = attributes[0];
    const matches = token === undefined
      ? current === undefined
      : current !== undefined && parseRaw(current.value) === Number(token);
    if (!matches) updates.push([attributeName, token]);
  }
  if (updates.length === 0) return false;

  let updated = original;
  for (const [attributeName, token] of updates) {
    updated = updateAttribute(updated, attributeName, token, partUri);
  }
  xml.replaceElement(propertiesElement, updated);
  return true;
}

function updateAttribute(
  template: string,
  name: string,
  value: string | undefined,
  partUri: string,
): string {
  const properties = LosslessXmlDocument.parse(template);
  const root = properties.roots[0];
  if (!root || root.localName !== 'tcPr') {
    throw new ModelParseError('Invalid table cell properties template', partUri);
  }
  const attributes = root.attributes.filter((attribute) => attribute.name === name);
  if (attributes.length > 1) {
    throw new ModelParseError(
      `Table cell contains repeated direct ${name} attributes`,
      partUri,
    );
  }
  const attribute = attributes[0];
  if (value !== undefined) {
    if (attribute) properties.replaceAttribute(attribute, value);
    else {
      const insertionPoint = root.selfClosing
        ? properties.source.lastIndexOf('/', root.startTagEnd - 1)
        : root.startTagEnd - 1;
      properties.replace(insertionPoint, insertionPoint, ` ${name}="${value}"`);
    }
  } else if (attribute) {
    let start = attribute.start;
    while (start > root.start && /[\t ]/.test(properties.source[start - 1] ?? '')) start -= 1;
    properties.replace(start, attribute.end, '');
  }
  return properties.serialize();
}

function parseRaw(value: string): number | undefined {
  if (!/^-?\d+$/.test(value)) return undefined;
  const raw = Number(value);
  return Number.isSafeInteger(raw) && raw >= MIN_INT32 && raw <= MAX_INT32
    ? raw
    : undefined;
}

function directChildren(element: XmlElement, localName: string): XmlElement[] {
  return element.children.filter(
    (child): child is XmlElement => child.type === 'element' && child.localName === localName,
  );
}
