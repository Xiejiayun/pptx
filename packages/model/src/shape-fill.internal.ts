import { type LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
import type { ShapeFill } from './preset-shape.js';
import {
  readSimpleFillChoice,
  renderSimpleFill,
  SIMPLE_FILL_CHOICE_NAMES,
  simpleFillsEqual,
} from './simple-fill.internal.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const FILL_CHOICES = new Set<string>(SIMPLE_FILL_CHOICE_NAMES);
const GEOMETRY_CHOICES = new Set(['prstGeom', 'custGeom']);

export function readShapeFill(
  _xml: LosslessXmlDocument,
  shape: XmlElement,
): ShapeFill | undefined {
  const properties = resolveShapeProperties(shape);
  if (!properties) return undefined;
  const choices = directChildren(properties).filter(isDrawingFillChoice);
  if (choices.length !== 1) return undefined;
  const choice = choices[0]!;
  if (!subtreeUsesDrawingNamespace(choice)) return undefined;
  return readSimpleFillChoice(choice, renderPrefix(choice));
}

export function replaceShapeFill(
  xml: LosslessXmlDocument,
  shape: XmlElement,
  fill: ShapeFill | undefined,
  partUri: string,
): boolean {
  const properties = resolveShapeProperties(shape);
  if (!properties) {
    throw new ModelParseError(
      'Shape must contain exactly one direct shape properties element',
      partUri,
    );
  }
  const choices = directChildren(properties).filter(isDrawingFillChoice);
  if (choices.length > 1) {
    throw new ModelParseError('Shape contains multiple direct fill choices', partUri);
  }

  const choice = choices[0];
  if (!choice && fill === undefined) return false;
  if (
    choice
    && fill !== undefined
    && subtreeUsesDrawingNamespace(choice)
    && simpleFillsEqual(
      readSimpleFillChoice(choice, renderPrefix(choice)),
      fill,
    )
  ) return false;

  if (fill === undefined) {
    xml.removeElement(choice!);
    return true;
  }

  if (choice) {
    const encoded = renderForParent(fill, choice, properties);
    xml.replaceElement(choice, encoded);
    return true;
  }

  const geometries = directChildren(properties).filter(
    (child) => GEOMETRY_CHOICES.has(child.localName)
      && namespaceUri(child) === DRAWING_NAMESPACE,
  );
  if (geometries.length !== 1) {
    throw new ModelParseError(
      'Shape must contain exactly one direct geometry before inserting a fill',
      partUri,
    );
  }
  const geometry = geometries[0]!;
  const encoded = renderForParent(fill, geometry, properties);
  xml.replace(geometry.end, geometry.end, encoded);
  return true;
}

function resolveShapeProperties(shape: XmlElement): XmlElement | undefined {
  if (shape.localName !== 'sp' || namespaceUri(shape) !== PRESENTATION_NAMESPACE) {
    return undefined;
  }
  const candidates = directChildren(shape).filter(
    ({ localName }) => localName === 'spPr',
  );
  if (candidates.length !== 1) return undefined;
  const properties = candidates[0]!;
  return namespaceUri(properties) === PRESENTATION_NAMESPACE
    ? properties
    : undefined;
}

function isDrawingFillChoice(element: XmlElement): boolean {
  return FILL_CHOICES.has(element.localName)
    && namespaceUri(element) === DRAWING_NAMESPACE;
}

function subtreeUsesDrawingNamespace(element: XmlElement): boolean {
  if (namespaceUri(element) !== DRAWING_NAMESPACE) return false;
  return directChildren(element).every(subtreeUsesDrawingNamespace);
}

function renderForParent(
  fill: ShapeFill,
  prefixSource: XmlElement,
  parent: XmlElement,
): string {
  const prefix = lexicalPrefix(prefixSource.name);
  const qualifiedPrefix = prefix === '' ? '' : `${prefix}:`;
  const encoded = renderSimpleFill(fill, qualifiedPrefix);
  if (namespaceUriForPrefix(parent, prefix) === DRAWING_NAMESPACE) return encoded;
  const tag = fill.kind === 'none' ? 'noFill' : 'solidFill';
  const declaration = prefix === ''
    ? ` xmlns="${DRAWING_NAMESPACE}"`
    : ` xmlns:${prefix}="${DRAWING_NAMESPACE}"`;
  return encoded.replace(`<${qualifiedPrefix}${tag}`, `<${qualifiedPrefix}${tag}${declaration}`);
}

function renderPrefix(element: XmlElement): string {
  const prefix = lexicalPrefix(element.name);
  return prefix === '' ? '' : `${prefix}:`;
}

function namespaceUri(element: XmlElement): string | undefined {
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

function directChildren(element: XmlElement): XmlElement[] {
  return element.children.filter(
    (child): child is XmlElement => child.type === 'element',
  );
}
