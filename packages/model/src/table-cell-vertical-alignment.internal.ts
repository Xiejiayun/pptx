import {
  escapeXmlAttribute,
  LosslessXmlDocument,
  type XmlElement,
} from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
import type { TextBoxVerticalAlignment } from './text.js';

const FROM_OOXML = new Map<string, TextBoxVerticalAlignment>([
  ['t', 'top'],
  ['ctr', 'middle'],
  ['b', 'bottom'],
]);

const TO_OOXML: Readonly<Record<TextBoxVerticalAlignment, string>> = {
  top: 't',
  middle: 'ctr',
  bottom: 'b',
};

export function renderTableCellVerticalAlignmentAttribute(
  value: TextBoxVerticalAlignment | undefined,
): string {
  return value === undefined ? '' : ` anchor="${TO_OOXML[value]}"`;
}

export function readTableCellVerticalAlignment(
  _xml: LosslessXmlDocument,
  cell: XmlElement,
): TextBoxVerticalAlignment | undefined {
  const properties = directChildren(cell, 'tcPr');
  if (properties.length !== 1) return undefined;
  const attributes = properties[0]!.attributes.filter(({ name }) => name === 'anchor');
  if (attributes.length !== 1) return undefined;
  return FROM_OOXML.get(attributes[0]!.value);
}

export function readTableVerticalAlignment(
  xml: LosslessXmlDocument,
  frame: XmlElement,
): TextBoxVerticalAlignment | undefined {
  const cells = physicalTableCells(frame);
  if (!cells) return undefined;
  const first = readTableCellVerticalAlignment(xml, cells[0]!);
  if (first === undefined) return undefined;
  return cells.every(
    (cell) => readTableCellVerticalAlignment(xml, cell) === first,
  ) ? first : undefined;
}

export function replaceTableCellVerticalAlignment(
  xml: LosslessXmlDocument,
  cell: XmlElement,
  value: TextBoxVerticalAlignment | undefined,
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
  const properties = LosslessXmlDocument.parse(xml.original(propertiesElement));
  const root = properties.roots[0];
  if (!root || root.localName !== 'tcPr') {
    throw new ModelParseError('Invalid table cell properties template', partUri);
  }
  const attributes = root.attributes.filter(({ name }) => name === 'anchor');
  if (attributes.length > 1) {
    throw new ModelParseError(
      'Table cell contains repeated direct vertical alignment attributes',
      partUri,
    );
  }
  const attribute = attributes[0];
  const token = value === undefined ? undefined : TO_OOXML[value];
  if (attribute?.value === token || (!attribute && token === undefined)) return false;

  if (token !== undefined) {
    if (attribute) properties.replaceAttribute(attribute, token);
    else {
      const insertionPoint = root.selfClosing
        ? properties.source.lastIndexOf('/', root.startTagEnd - 1)
        : root.startTagEnd - 1;
      properties.replace(
        insertionPoint,
        insertionPoint,
        ` anchor="${escapeXmlAttribute(token)}"`,
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

export function replaceTableVerticalAlignment(
  xml: LosslessXmlDocument,
  frame: XmlElement,
  value: TextBoxVerticalAlignment | undefined,
  partUri: string,
): boolean {
  const cells = physicalTableCells(frame);
  if (!cells) {
    throw new ModelParseError(
      'Table must contain one complete set of direct physical cells',
      partUri,
    );
  }
  let changed = false;
  for (const cell of cells) {
    changed = replaceTableCellVerticalAlignment(
      xml,
      cell,
      value,
      partUri,
    ) || changed;
  }
  return changed;
}

function physicalTableCells(frame: XmlElement): readonly XmlElement[] | undefined {
  if (frame.localName !== 'graphicFrame') return undefined;
  const graphic = exactDirectChild(frame, 'graphic');
  const graphicData = graphic ? exactDirectChild(graphic, 'graphicData') : undefined;
  const table = graphicData ? exactDirectChild(graphicData, 'tbl') : undefined;
  if (!table) return undefined;
  const rows = directChildren(table, 'tr');
  if (rows.length === 0) return undefined;
  const matrix = rows.map((row) => directChildren(row, 'tc'));
  return matrix.some((cells) => cells.length === 0) ? undefined : matrix.flat();
}

function exactDirectChild(
  element: XmlElement,
  localName: string,
): XmlElement | undefined {
  const matches = directChildren(element, localName);
  return matches.length === 1 ? matches[0] : undefined;
}

function directChildren(element: XmlElement, localName: string): XmlElement[] {
  return element.children.filter(
    (child): child is XmlElement => child.type === 'element' && child.localName === localName,
  );
}
