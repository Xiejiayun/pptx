import {
  escapeXmlAttribute,
  LosslessXmlDocument,
  type XmlElement,
} from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
import type { TableCellTextDirection } from './shapes.js';
import { readDirectTablePhysicalCells } from './table-physical-cells.internal.js';

const DIRECTIONS = new Set<TableCellTextDirection>([
  'horz',
  'vert',
  'vert270',
  'wordArtVert',
]);

export function readTableCellTextDirection(
  _xml: LosslessXmlDocument,
  cell: XmlElement,
): TableCellTextDirection | undefined {
  const properties = directChildren(cell, 'tcPr');
  if (properties.length !== 1) return undefined;
  const attributes = properties[0]!.attributes.filter(({ name }) => name === 'vert');
  if (attributes.length !== 1) return undefined;
  const value = attributes[0]!.value;
  return DIRECTIONS.has(value as TableCellTextDirection)
    ? value as TableCellTextDirection
    : undefined;
}

export function readTableTextDirection(
  xml: LosslessXmlDocument,
  frame: XmlElement,
): TableCellTextDirection | undefined {
  const cells = readDirectTablePhysicalCells(frame);
  if (!cells) return undefined;
  const first = readTableCellTextDirection(xml, cells[0]!);
  if (first === undefined) return undefined;
  return cells.every(
    (cell) => readTableCellTextDirection(xml, cell) === first,
  ) ? first : undefined;
}

export function normalizeTableCellTextDirection(
  value: unknown,
  context: string,
): TableCellTextDirection {
  if (typeof value !== 'string' || !DIRECTIONS.has(value as TableCellTextDirection)) {
    throw new TypeError(`${context} must be horz, vert, vert270, or wordArtVert`);
  }
  return value as TableCellTextDirection;
}

export function renderTableCellTextDirectionAttribute(
  value: TableCellTextDirection | undefined,
): string {
  return value === undefined || value === 'horz'
    ? ''
    : ` vert="${escapeXmlAttribute(value)}"`;
}

export function replaceTableCellTextDirection(
  xml: LosslessXmlDocument,
  cell: XmlElement,
  value: TableCellTextDirection | undefined,
  partUri: string,
): boolean {
  const directProperties = directChildren(cell, 'tcPr');
  if (directProperties.length !== 1) {
    throw new ModelParseError('Table cell must contain exactly one direct cell properties element', partUri);
  }

  const propertiesElement = directProperties[0]!;
  const properties = LosslessXmlDocument.parse(xml.original(propertiesElement));
  const root = properties.roots[0];
  if (!root || root.localName !== 'tcPr') {
    throw new ModelParseError('Invalid table cell properties template', partUri);
  }
  const attributes = root.attributes.filter(({ name }) => name === 'vert');
  if (attributes.length > 1) {
    throw new ModelParseError('Table cell contains repeated direct text direction attributes', partUri);
  }
  const attribute = attributes[0];
  if (attribute?.value === value || (!attribute && value === undefined)) return false;

  if (value !== undefined) {
    if (attribute) properties.replaceAttribute(attribute, value);
    else {
      const insertionPoint = root.selfClosing
        ? properties.source.lastIndexOf('/', root.startTagEnd - 1)
        : root.startTagEnd - 1;
      properties.replace(
        insertionPoint,
        insertionPoint,
        ` vert="${escapeXmlAttribute(value)}"`,
      );
    }
  } else if (attribute) {
    let start = attribute.start;
    while (start > root.start && /[\t ]/.test(properties.source[start - 1] ?? '')) start -= 1;
    properties.replace(start, attribute.end, '');
  }

  xml.replaceElement(propertiesElement, properties.serialize());
  return true;
}

export function replaceTableTextDirection(
  xml: LosslessXmlDocument,
  frame: XmlElement,
  value: TableCellTextDirection | undefined,
  partUri: string,
): boolean {
  const cells = readDirectTablePhysicalCells(frame);
  if (!cells) {
    throw new ModelParseError(
      'Table must contain one complete set of direct physical cells',
      partUri,
    );
  }
  let changed = false;
  for (const cell of cells) {
    changed = replaceTableCellTextDirection(xml, cell, value, partUri) || changed;
  }
  return changed;
}

function directChildren(element: XmlElement, localName: string): XmlElement[] {
  return element.children.filter(
    (child): child is XmlElement => child.type === 'element' && child.localName === localName,
  );
}
