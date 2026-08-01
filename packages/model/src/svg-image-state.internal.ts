import type {
  LosslessXmlDocument,
  XmlAttribute,
  XmlElement,
} from '@pptx/lossless-xml';
import type { OpcPackage, Relationship } from '@pptx/opc';
import {
  SVG_IMAGE_EXTENSION_URI,
  SVG_IMAGE_NAMESPACE,
} from './svg-image-create.internal.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const IMAGE_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

export interface SvgImageState {
  readonly fallbackReference: XmlAttribute;
  readonly svgReference: XmlAttribute;
  readonly fallbackRelationship: Relationship;
  readonly svgRelationship: Relationship;
  readonly fallbackPartUri: string;
  readonly svgPartUri: string;
}

export function readSvgImageState(
  xml: LosslessXmlDocument,
  picture: XmlElement,
  relationships: readonly Relationship[],
  pkg: OpcPackage,
): SvgImageState | undefined {
  if (!isElement(picture, 'pic', PRESENTATION_NAMESPACE)) return undefined;
  const fills = directChildrenWithLocalName(picture, 'blipFill');
  if (fills.length !== 1 || elementNamespaceUri(fills[0]!) !== PRESENTATION_NAMESPACE) {
    return undefined;
  }
  const blips = directChildrenWithLocalName(fills[0]!, 'blip');
  if (blips.length !== 1 || elementNamespaceUri(blips[0]!) !== DRAWING_NAMESPACE) {
    return undefined;
  }
  const blip = blips[0]!;
  const fallbackReference = relationshipEmbedAttribute(blip);
  if (!fallbackReference) return undefined;

  const extensionLists = directChildrenWithLocalName(blip, 'extLst');
  if (
    extensionLists.length !== 1
    || elementNamespaceUri(extensionLists[0]!) !== DRAWING_NAMESPACE
  ) return undefined;
  const extensions = directChildren(extensionLists[0]!).filter((element) =>
    element.localName === 'ext'
    && unqualifiedAttributeValue(element, 'uri') === SVG_IMAGE_EXTENSION_URI);
  if (extensions.length !== 1 || elementNamespaceUri(extensions[0]!) !== DRAWING_NAMESPACE) {
    return undefined;
  }
  const extension = extensions[0]!;
  const extensionChildren = directChildren(extension);
  if (
    extensionChildren.length !== 1
    || hasNonWhitespaceText(extension)
    || !isElement(extensionChildren[0]!, 'svgBlip', SVG_IMAGE_NAMESPACE)
  ) return undefined;
  const svgBlip = extensionChildren[0]!;
  if (directChildren(svgBlip).length > 0 || hasNonWhitespaceText(svgBlip)) return undefined;
  const svgReference = relationshipEmbedAttribute(svgBlip);
  if (!svgReference || hasUnsupportedSvgBlipAttribute(svgBlip, svgReference)) {
    return undefined;
  }

  const fallbackRelationships = relationships.filter(
    ({ id }) => id === fallbackReference.value,
  );
  const svgRelationships = relationships.filter(({ id }) => id === svgReference.value);
  if (fallbackRelationships.length !== 1 || svgRelationships.length !== 1) {
    return undefined;
  }
  const fallbackRelationship = fallbackRelationships[0];
  const svgRelationship = svgRelationships[0];
  if (
    !isInternalImageRelationship(fallbackRelationship, pkg)
    || !isInternalImageRelationship(svgRelationship, pkg)
    || fallbackRelationship.id === svgRelationship.id
    || fallbackRelationship.resolvedTarget === svgRelationship.resolvedTarget
  ) return undefined;

  return Object.freeze({
    fallbackReference,
    svgReference,
    fallbackRelationship,
    svgRelationship,
    fallbackPartUri: fallbackRelationship.resolvedTarget,
    svgPartUri: svgRelationship.resolvedTarget,
  });
}

export function hasSvgImageExtensionCandidate(
  xml: LosslessXmlDocument,
  picture: XmlElement,
): boolean {
  return xml.descendants(picture).some((element) =>
    (
      element.localName === 'svgBlip'
      && elementNamespaceUri(element) === SVG_IMAGE_NAMESPACE
    )
    || (
      element.localName === 'ext'
      && unqualifiedAttributeValue(element, 'uri') === SVG_IMAGE_EXTENSION_URI
    ));
}

export function relationshipReferenceCount(
  xml: LosslessXmlDocument,
  id: string,
): number {
  let count = 0;
  for (const element of xml.elements()) {
    for (const attribute of element.attributes) {
      if (
        attribute.value === id
        && attributeNamespaceUri(element, attribute) === RELATIONSHIP_NAMESPACE
      ) count += 1;
    }
  }
  return count;
}

export function elementNamespaceUri(element: XmlElement): string | undefined {
  return namespaceUriForPrefix(element, lexicalPrefix(element.name));
}

export function attributeNamespaceUri(
  element: XmlElement,
  attribute: XmlAttribute,
): string | undefined {
  const prefix = lexicalPrefix(attribute.name);
  if (prefix === 'xmlns' || attribute.name === 'xmlns') {
    return 'http://www.w3.org/2000/xmlns/';
  }
  if (prefix.length === 0) return undefined;
  if (prefix === 'xml') return 'http://www.w3.org/XML/1998/namespace';
  return namespaceUriForPrefix(element, prefix);
}

function relationshipEmbedAttribute(element: XmlElement): XmlAttribute | undefined {
  const embedAttributes = element.attributes.filter(({ localName }) => localName === 'embed');
  if (
    embedAttributes.length !== 1
    || attributeNamespaceUri(element, embedAttributes[0]!) !== RELATIONSHIP_NAMESPACE
  ) return undefined;
  if (element.attributes.some((attribute) =>
    attribute.localName === 'link'
    && attributeNamespaceUri(element, attribute) === RELATIONSHIP_NAMESPACE)) {
    return undefined;
  }
  return embedAttributes[0]!;
}

function isInternalImageRelationship(
  relationship: Relationship | undefined,
  pkg: OpcPackage,
): relationship is Relationship & { readonly resolvedTarget: string } {
  return relationship?.type === IMAGE_RELATIONSHIP
    && relationship.targetMode === 'Internal'
    && typeof relationship.resolvedTarget === 'string'
    && pkg.hasPart(relationship.resolvedTarget);
}

function hasUnsupportedSvgBlipAttribute(
  element: XmlElement,
  reference: XmlAttribute,
): boolean {
  return element.attributes.some((attribute) =>
    attribute !== reference
    && attribute.name !== 'xmlns'
    && !attribute.name.startsWith('xmlns:'));
}

function directChildren(element: XmlElement): XmlElement[] {
  return element.children.filter((child): child is XmlElement => child.type === 'element');
}

function directChildrenWithLocalName(
  element: XmlElement,
  localName: string,
): XmlElement[] {
  return directChildren(element).filter((child) => child.localName === localName);
}

function isElement(
  element: XmlElement,
  localName: string,
  namespace: string,
): boolean {
  return element.localName === localName && elementNamespaceUri(element) === namespace;
}

function unqualifiedAttributeValue(
  element: XmlElement,
  name: string,
): string | undefined {
  return element.attributes.find((attribute) =>
    attribute.name === name && attribute.localName === name)?.value;
}

function hasNonWhitespaceText(element: XmlElement): boolean {
  return element.children.some((child) => child.type === 'text' && /\S/.test(child.value));
}

function namespaceUriForPrefix(
  element: XmlElement,
  prefix: string,
): string | undefined {
  const declarationName = prefix.length === 0 ? 'xmlns' : `xmlns:${prefix}`;
  let current: XmlElement | undefined = element;
  while (current) {
    const declaration = current.attributes.find(({ name }) => name === declarationName);
    if (declaration) return declaration.value.length === 0 ? undefined : declaration.value;
    current = current.parent;
  }
  return undefined;
}

function lexicalPrefix(name: string): string {
  const separator = name.indexOf(':');
  return separator < 0 ? '' : name.slice(0, separator);
}
