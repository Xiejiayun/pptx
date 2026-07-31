import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
import {
  normalizeSimpleFill,
  readSimpleFillChoice,
  renderSimpleFill,
  SIMPLE_FILL_CHOICE_NAMES,
  simpleFillsEqual,
} from './simple-fill.internal.js';
import type { TableCellFill } from './shapes.js';

const FILL_CHOICES = new Set<string>(SIMPLE_FILL_CHOICE_NAMES);

export function normalizeTableCellFill(
  value: unknown,
  context: string,
): TableCellFill | undefined {
  return normalizeSimpleFill(value, context);
}

export function readTableCellFill(
  _xml: LosslessXmlDocument,
  cell: XmlElement,
): TableCellFill | undefined {
  const directProperties = directChildren(cell).filter(({ localName }) => localName === 'tcPr');
  if (directProperties.length !== 1) return undefined;
  const properties = directProperties[0]!;
  const prefix = lexicalPrefix(properties.name);
  const choices = directChildren(properties).filter(
    (child) => FILL_CHOICES.has(child.localName) && child.name === `${prefix}${child.localName}`,
  );
  if (choices.length !== 1) return undefined;
  return readSimpleFillChoice(choices[0]!, prefix);
}

export function replaceTableCellFill(
  xml: LosslessXmlDocument,
  cell: XmlElement,
  fill: TableCellFill | undefined,
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
  const properties = LosslessXmlDocument.parse(xml.original(propertiesElement));
  const root = properties.roots[0];
  if (!root || root.localName !== 'tcPr') {
    throw new ModelParseError('Invalid table cell properties template', partUri);
  }
  const prefix = lexicalPrefix(root.name);
  const choices = directChildren(root).filter(
    (child) => FILL_CHOICES.has(child.localName) && child.name === `${prefix}${child.localName}`,
  );
  if (choices.length > 1) {
    throw new ModelParseError('Table cell contains multiple direct fill choices', partUri);
  }

  const choice = choices[0];
  if (!choice && fill === undefined) return false;
  if (choice && fill !== undefined && simpleFillsEqual(readSimpleFillChoice(choice, prefix), fill)) {
    return false;
  }

  if (fill === undefined) {
    if (choice) properties.removeElement(choice);
  } else {
    const encoded = renderTableCellFill(fill, prefix);
    if (choice) properties.replaceElement(choice, encoded);
    else {
      const extension = directChildren(root).find(
        (child) => child.name === `${prefix}extLst`,
      );
      if (extension) properties.replace(extension.start, extension.start, encoded);
      else properties.appendChildXml(root, encoded);
    }
  }

  xml.replaceElement(propertiesElement, properties.serialize());
  return true;
}

export function renderTableCellFill(fill: TableCellFill, prefix: string): string {
  return renderSimpleFill(fill, prefix);
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
