import type {
  LosslessXmlDocument,
  XmlAttribute,
  XmlElement,
} from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const TRUE_TOKENS = new Set(['1', 'true', 'on']);
const FALSE_TOKENS = new Set(['0', 'false', 'off']);

interface TextShapeIsTextBoxState {
  readonly properties: XmlElement;
  readonly attribute?: XmlAttribute;
}

export function readTextShapeIsTextBox(
  _xml: LosslessXmlDocument,
  shape: XmlElement,
): boolean | undefined {
  const state = resolveState(shape);
  if (!state) return undefined;
  if (!state.attribute) return false;
  if (TRUE_TOKENS.has(state.attribute.value)) return true;
  if (FALSE_TOKENS.has(state.attribute.value)) return false;
  return undefined;
}

export function replaceTextShapeIsTextBox(
  xml: LosslessXmlDocument,
  shape: XmlElement,
  value: boolean,
  partUri: string,
): boolean {
  const state = resolveState(shape);
  if (!state) {
    throw new ModelParseError('Text shape isTextBox state is not safely editable', partUri);
  }

  if (value) {
    if (state.attribute?.value === '1') return false;
    if (state.attribute) {
      xml.replaceAttribute(state.attribute, '1');
      return true;
    }
    insertAttribute(xml, state.properties, ' txBox="1"', partUri);
    return true;
  }

  if (!state.attribute) return false;
  removeAttribute(xml, state.properties, state.attribute);
  return true;
}

function resolveState(shape: XmlElement): TextShapeIsTextBoxState | undefined {
  if (
    shape.localName !== 'sp'
    || elementNamespaceUri(shape) !== PRESENTATION_NAMESPACE
  ) return undefined;

  const nonVisualCandidates = directChildren(shape).filter(
    ({ localName }) => localName === 'nvSpPr',
  );
  if (
    nonVisualCandidates.length !== 1
    || elementNamespaceUri(nonVisualCandidates[0]!) !== PRESENTATION_NAMESPACE
  ) return undefined;

  const propertyCandidates = directChildren(nonVisualCandidates[0]!).filter(
    ({ localName }) => localName === 'cNvSpPr',
  );
  if (
    propertyCandidates.length !== 1
    || elementNamespaceUri(propertyCandidates[0]!) !== PRESENTATION_NAMESPACE
  ) return undefined;

  const properties = propertyCandidates[0]!;
  const attributes = properties.attributes.filter((attribute) =>
    attribute.localName === 'txBox'
    && attribute.name !== 'xmlns'
    && !attribute.name.startsWith('xmlns:'));
  if (attributes.length > 1 || (attributes[0] && attributes[0].name !== 'txBox')) {
    return undefined;
  }
  return attributes[0]
    ? { properties, attribute: attributes[0] }
    : { properties };
}

function insertAttribute(
  xml: LosslessXmlDocument,
  element: XmlElement,
  attribute: string,
  partUri: string,
): void {
  const position = element.startTagEnd - (element.selfClosing ? 2 : 1);
  if (position <= element.start) {
    throw new ModelParseError('Text shape isTextBox state is not safely editable', partUri);
  }
  xml.replace(position, position, attribute);
}

function removeAttribute(
  xml: LosslessXmlDocument,
  element: XmlElement,
  attribute: XmlAttribute,
): void {
  let start = attribute.start;
  while (
    start > element.start
    && (xml.source[start - 1] === ' ' || xml.source[start - 1] === '\t')
  ) start -= 1;
  xml.replace(start, attribute.end, '');
}

function directChildren(element: XmlElement): XmlElement[] {
  return element.children.filter(
    (child): child is XmlElement => child.type === 'element',
  );
}

function elementNamespaceUri(element: XmlElement): string | undefined {
  return namespaceUriForPrefix(element, lexicalPrefix(element.name));
}

function namespaceUriForPrefix(
  element: XmlElement,
  prefix: string,
): string | undefined {
  const declarationName = prefix === '' ? 'xmlns' : `xmlns:${prefix}`;
  for (let current: XmlElement | undefined = element; current; current = current.parent) {
    const declarations = current.attributes.filter(({ name }) => name === declarationName);
    if (declarations.length > 1) return undefined;
    if (declarations[0]) return declarations[0].value;
  }
  return undefined;
}

function lexicalPrefix(name: string): string {
  const separator = name.indexOf(':');
  return separator < 0 ? '' : name.slice(0, separator);
}
