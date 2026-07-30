import {
  LosslessXmlDocument,
  type XmlAttribute,
  type XmlElement,
} from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const TRUE_TOKENS = new Set(['1', 'true', 'on']);
const FALSE_TOKENS = new Set(['0', 'false', 'off']);

interface SlideVisibilityState {
  readonly root: XmlElement;
  readonly show?: XmlAttribute;
}

export function readSlideHidden(
  xml: LosslessXmlDocument,
): boolean | undefined {
  const state = resolveSlideVisibilityState(xml);
  if (!state) return undefined;
  if (!state.show) return false;
  if (FALSE_TOKENS.has(state.show.value)) return true;
  if (TRUE_TOKENS.has(state.show.value)) return false;
  return undefined;
}

export function replaceSlideHidden(
  xml: LosslessXmlDocument,
  value: boolean,
): boolean {
  const state = resolveSlideVisibilityState(xml);
  if (!state) {
    throw new ModelParseError('Slide visibility is not safely editable');
  }
  if (value) {
    if (state.show?.value === '0') return false;
    if (state.show) xml.replaceAttribute(state.show, '0');
    else insertRootAttribute(xml, state.root, ' show="0"');
    return true;
  }
  if (!state.show) return false;
  removeRootAttribute(xml, state.root, state.show);
  return true;
}

function resolveSlideVisibilityState(
  xml: LosslessXmlDocument,
): SlideVisibilityState | undefined {
  if (xml.roots.length !== 1) return undefined;
  const root = xml.roots[0];
  if (
    !root
    || root.localName !== 'sld'
    || namespaceUriForPrefix(root, lexicalPrefix(root.name)) !== PRESENTATION_NAMESPACE
  ) return undefined;

  const showAttributes = root.attributes.filter(({ name }) => name === 'show');
  if (showAttributes.length > 1) return undefined;
  const show = showAttributes[0];
  return show ? { root, show } : { root };
}

function namespaceUriForPrefix(
  element: XmlElement,
  prefix: string,
): string | undefined {
  let current: XmlElement | undefined = element;
  const declarationName = prefix === '' ? 'xmlns' : `xmlns:${prefix}`;
  while (current) {
    const declarations = current.attributes.filter(
      ({ name }) => name === declarationName,
    );
    if (declarations.length > 1) return undefined;
    if (declarations[0]) return declarations[0].value;
    current = current.parent;
  }
  return undefined;
}

function lexicalPrefix(name: string): string {
  const separator = name.indexOf(':');
  return separator < 0 ? '' : name.slice(0, separator);
}

function insertRootAttribute(
  xml: LosslessXmlDocument,
  root: XmlElement,
  attribute: string,
): void {
  const position = root.selfClosing
    ? xml.source.lastIndexOf('/', root.startTagEnd - 1)
    : root.startTagEnd - 1;
  if (position <= root.start) {
    throw new ModelParseError('Slide visibility is not safely editable');
  }
  xml.replace(position, position, attribute);
}

function removeRootAttribute(
  xml: LosslessXmlDocument,
  root: XmlElement,
  attribute: XmlAttribute,
): void {
  let start = attribute.start;
  while (
    start > root.start
    && (xml.source[start - 1] === ' ' || xml.source[start - 1] === '\t')
  ) start -= 1;
  xml.replace(start, attribute.end, '');
}
