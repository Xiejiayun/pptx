import {
  escapeXmlAttribute,
  LosslessXmlDocument,
  type XmlElement,
} from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
import type { TextAlignment } from './text.js';

const FROM_OOXML = new Map<string, TextAlignment>([
  ['l', 'left'],
  ['ctr', 'center'],
  ['r', 'right'],
  ['just', 'justify'],
]);

const TO_OOXML: Readonly<Record<TextAlignment, string>> = {
  left: 'l',
  center: 'ctr',
  right: 'r',
  justify: 'just',
};

export function readTableCellHorizontalAlignment(
  _xml: LosslessXmlDocument,
  cell: XmlElement,
): TextAlignment | undefined {
  const textBodies = directChildren(cell, 'txBody');
  if (textBodies.length !== 1) return undefined;
  const paragraphs = directChildren(textBodies[0]!, 'p');
  if (paragraphs.length !== 1) return undefined;
  const properties = directChildren(paragraphs[0]!, 'pPr');
  if (properties.length !== 1) return undefined;
  const attributes = properties[0]!.attributes.filter(({ name }) => name === 'algn');
  if (attributes.length !== 1) return undefined;
  return FROM_OOXML.get(attributes[0]!.value);
}

export function replaceTableCellHorizontalAlignment(
  xml: LosslessXmlDocument,
  cell: XmlElement,
  value: TextAlignment | undefined,
  partUri: string,
): boolean {
  const textBodies = directChildren(cell, 'txBody');
  if (textBodies.length !== 1) {
    throw new ModelParseError(
      'Table cell must contain exactly one direct text body',
      partUri,
    );
  }
  const paragraphs = directChildren(textBodies[0]!, 'p');
  if (paragraphs.length !== 1) {
    throw new ModelParseError(
      'Table cell must contain exactly one direct text paragraph',
      partUri,
    );
  }

  const paragraphElement = paragraphs[0]!;
  const paragraphXml = LosslessXmlDocument.parse(xml.original(paragraphElement));
  const root = paragraphXml.roots[0];
  if (!root || root.localName !== 'p' || root.selfClosing) {
    throw new ModelParseError('Invalid table cell paragraph template', partUri);
  }

  const properties = directChildren(root, 'pPr');
  if (properties.length > 1) {
    throw new ModelParseError(
      'Table cell paragraph contains repeated direct properties elements',
      partUri,
    );
  }

  const token = value === undefined ? undefined : TO_OOXML[value];
  const propertiesElement = properties[0];
  if (!propertiesElement) {
    if (token === undefined) return false;
    const firstElement = root.children.find(
      (child): child is XmlElement => child.type === 'element',
    );
    const insertionPoint = firstElement?.start ?? root.endTagStart;
    const separator = root.name.includes(':')
      ? `${root.name.slice(0, root.name.indexOf(':'))}:`
      : '';
    paragraphXml.replace(
      insertionPoint,
      insertionPoint,
      `<${separator}pPr algn="${escapeXmlAttribute(token)}"/>`,
    );
    xml.replaceElement(paragraphElement, paragraphXml.serialize());
    return true;
  }

  const attributes = propertiesElement.attributes.filter(({ name }) => name === 'algn');
  if (attributes.length > 1) {
    throw new ModelParseError(
      'Table cell paragraph contains repeated direct horizontal alignment attributes',
      partUri,
    );
  }
  const attribute = attributes[0];
  if (attribute?.value === token || (!attribute && token === undefined)) return false;

  if (token !== undefined) {
    if (attribute) paragraphXml.replaceAttribute(attribute, token);
    else {
      const insertionPoint = propertiesElement.selfClosing
        ? paragraphXml.source.lastIndexOf('/', propertiesElement.startTagEnd - 1)
        : propertiesElement.startTagEnd - 1;
      paragraphXml.replace(
        insertionPoint,
        insertionPoint,
        ` algn="${escapeXmlAttribute(token)}"`,
      );
    }
  } else if (attribute) {
    let start = attribute.start;
    while (start > propertiesElement.start && /[\t ]/.test(paragraphXml.source[start - 1] ?? '')) {
      start -= 1;
    }
    paragraphXml.replace(start, attribute.end, '');
  }

  xml.replaceElement(paragraphElement, paragraphXml.serialize());
  return true;
}

function directChildren(element: XmlElement, localName: string): XmlElement[] {
  return element.children.filter(
    (child): child is XmlElement => child.type === 'element' && child.localName === localName,
  );
}
