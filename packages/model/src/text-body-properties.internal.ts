import {
  escapeXmlAttribute,
  LosslessXmlDocument,
  type XmlElement,
} from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';

export function requireTextBodyProperties(
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

export function updateTextBodyAttribute(
  template: string,
  name: string,
  value: string | undefined,
): string {
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
