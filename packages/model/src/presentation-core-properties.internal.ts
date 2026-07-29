import {
  escapeXmlText,
  LosslessXmlDocument,
  type XmlElement,
} from '@pptx/lossless-xml';
import {
  OpcPackage,
  PackageError,
  type PackagePart,
  type Relationship,
} from '@pptx/opc';
import { ModelParseError } from './errors.js';

const CORE_PROPERTIES_RELATIONSHIP =
  'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';
const CORE_PROPERTIES_CONTENT_TYPE =
  'application/vnd.openxmlformats-package.core-properties+xml';
const CORE_PROPERTIES_NAMESPACE =
  'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';

export interface CoreTextPropertyDescriptor {
  readonly label: string;
  readonly localName: string;
  readonly namespace: string;
  readonly preferredPrefix: string;
}

interface CorePropertiesState {
  readonly relationship: Relationship;
  readonly part: PackagePart;
  readonly xml: LosslessXmlDocument;
  readonly root: XmlElement;
  readonly properties: readonly XmlElement[];
}

export function readCoreTextProperty(
  pkg: OpcPackage,
  descriptor: CoreTextPropertyDescriptor,
): string | undefined {
  const state = readCorePropertiesState(pkg, descriptor);
  if (!state || state.properties.length !== 1) return undefined;
  return readSimpleProperty(state.xml, state.properties[0]!);
}

export function replaceCoreTextProperty(
  pkg: OpcPackage,
  descriptor: CoreTextPropertyDescriptor,
  value: string | undefined,
): void {
  const relationships = corePropertiesRelationships(pkg);
  if (relationships.length === 0) {
    if (value === undefined) return;
    createCoreProperties(pkg, descriptor, value);
    return;
  }
  if (relationships.length !== 1) {
    throw new PackageError('Presentation has multiple core-properties relationships', '/');
  }

  const state = requireCorePropertiesState(pkg, relationships[0]!, descriptor);
  if (state.properties.length > 1) {
    throw new ModelParseError(
      `Core properties contain multiple direct ${descriptor.label}s`,
      state.part.uri,
    );
  }
  const property = state.properties[0];
  const current = property ? readSimpleProperty(state.xml, property) : undefined;
  if (property && current === undefined) {
    throw new ModelParseError(
      `Core properties ${descriptor.label} is not simple text`,
      state.part.uri,
    );
  }
  if (property && value !== undefined && current === value) return;
  if (!property && value === undefined) return;

  if (value === undefined) {
    state.xml.removeElement(property!);
  } else if (!property) {
    state.xml.appendChildXml(
      state.root,
      renderInsertedProperty(state.root, descriptor, value),
    );
  } else if (property.selfClosing) {
    state.xml.replaceElement(
      property,
      expandSelfClosingProperty(state.xml, property, descriptor, value),
    );
  } else {
    state.xml.replaceText(property, value);
  }
  pkg.setPart(state.part.uri, state.xml.serialize(), state.part.contentType);
}

function corePropertiesRelationships(pkg: OpcPackage): readonly Relationship[] {
  return pkg.relationships('/').filter(
    ({ type }) => type === CORE_PROPERTIES_RELATIONSHIP,
  );
}

function readCorePropertiesState(
  pkg: OpcPackage,
  descriptor: CoreTextPropertyDescriptor,
): CorePropertiesState | undefined {
  const relationships = corePropertiesRelationships(pkg);
  if (relationships.length !== 1) return undefined;
  const relationship = relationships[0]!;
  if (relationship.targetMode !== 'Internal' || !relationship.resolvedTarget) return undefined;
  const part = pkg.getPart(relationship.resolvedTarget);
  if (!part || part.contentType !== CORE_PROPERTIES_CONTENT_TYPE) return undefined;
  try {
    return parseCorePropertiesState(relationship, part, descriptor);
  } catch {
    return undefined;
  }
}

function requireCorePropertiesState(
  pkg: OpcPackage,
  relationship: Relationship,
  descriptor: CoreTextPropertyDescriptor,
): CorePropertiesState {
  if (relationship.targetMode !== 'Internal' || !relationship.resolvedTarget) {
    throw new PackageError('Core-properties relationship must be internal', '/');
  }
  const part = pkg.getPart(relationship.resolvedTarget);
  if (!part) {
    throw new PackageError(
      'Core-properties relationship target is missing',
      relationship.resolvedTarget,
    );
  }
  if (part.contentType !== CORE_PROPERTIES_CONTENT_TYPE) {
    throw new PackageError('Core-properties part has an unsupported content type', part.uri);
  }
  try {
    return parseCorePropertiesState(relationship, part, descriptor);
  } catch (error) {
    if (error instanceof ModelParseError) throw error;
    throw new ModelParseError(
      error instanceof Error ? error.message : String(error),
      part.uri,
    );
  }
}

function parseCorePropertiesState(
  relationship: Relationship,
  part: PackagePart,
  descriptor: CoreTextPropertyDescriptor,
): CorePropertiesState {
  const xml = LosslessXmlDocument.parse(part.bytes);
  if (xml.roots.length !== 1) {
    throw new ModelParseError('Core-properties part must have one root', part.uri);
  }
  const root = xml.roots[0]!;
  if (
    root.localName !== 'coreProperties'
    || namespaceUri(root) !== CORE_PROPERTIES_NAMESPACE
  ) {
    throw new ModelParseError('Core-properties root is invalid', part.uri);
  }
  const properties = directChildren(root).filter(
    (child) => child.localName === descriptor.localName
      && namespaceUri(child) === descriptor.namespace,
  );
  return { relationship, part, xml, root, properties };
}

function readSimpleProperty(
  xml: LosslessXmlDocument,
  property: XmlElement,
): string | undefined {
  if (directChildren(property).length > 0) return undefined;
  if (/<!\[CDATA\[/i.test(xml.original(property))) return undefined;
  return xml.text(property);
}

function createCoreProperties(
  pkg: OpcPackage,
  descriptor: CoreTextPropertyDescriptor,
  value: string,
): void {
  const canonicalUri = '/docProps/core.xml';
  const partUri = pkg.hasPart(canonicalUri)
    ? pkg.allocatePartUri('/docProps', 'core', '.xml')
    : canonicalUri;
  const qualifiedName = `${descriptor.preferredPrefix}:${descriptor.localName}`;
  const preferredNamespaceDeclaration =
    descriptor.preferredPrefix === 'cp' && descriptor.namespace === CORE_PROPERTIES_NAMESPACE
      ? ''
      : ` xmlns:${descriptor.preferredPrefix}="${descriptor.namespace}"`;
  const xml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`
    + `<cp:coreProperties xmlns:cp="${CORE_PROPERTIES_NAMESPACE}"${preferredNamespaceDeclaration}>`
    + `<${qualifiedName}>${escapeXmlText(value)}</${qualifiedName}>`
    + '</cp:coreProperties>';
  pkg.setPart(partUri, xml, CORE_PROPERTIES_CONTENT_TYPE);
  pkg.addRelationship('/', {
    type: CORE_PROPERTIES_RELATIONSHIP,
    target: partUri.slice(1),
  });
}

function renderInsertedProperty(
  root: XmlElement,
  descriptor: CoreTextPropertyDescriptor,
  value: string,
): string {
  const prefix = prefixForNamespace(root, descriptor.namespace);
  if (prefix === undefined) {
    const qualifiedName = `${descriptor.preferredPrefix}:${descriptor.localName}`;
    return `<${qualifiedName} xmlns:${descriptor.preferredPrefix}="${descriptor.namespace}">`
      + `${escapeXmlText(value)}</${qualifiedName}>`;
  }
  const name = prefix === '' ? descriptor.localName : `${prefix}:${descriptor.localName}`;
  return `<${name}>${escapeXmlText(value)}</${name}>`;
}

function expandSelfClosingProperty(
  xml: LosslessXmlDocument,
  property: XmlElement,
  descriptor: CoreTextPropertyDescriptor,
  value: string,
): string {
  const original = xml.original(property);
  const marker = original.lastIndexOf('/>');
  if (marker < 0) {
    throw new ModelParseError(
      `Self-closing core-properties ${descriptor.label} is malformed`,
    );
  }
  const open = `${original.slice(0, marker).replace(/\s+$/, '')}>`;
  return `${open}${escapeXmlText(value)}</${property.name}>`;
}

function namespaceUri(element: XmlElement): string | undefined {
  const prefix = lexicalPrefix(element.name);
  const declaration = prefix === '' ? 'xmlns' : `xmlns:${prefix}`;
  for (let scope: XmlElement | undefined = element; scope; scope = scope.parent) {
    const attribute = scope.attributes.find(({ name }) => name === declaration);
    if (attribute) return attribute.value;
  }
  return undefined;
}

function prefixForNamespace(element: XmlElement, uri: string): string | undefined {
  for (let scope: XmlElement | undefined = element; scope; scope = scope.parent) {
    for (const attribute of scope.attributes) {
      if (attribute.value !== uri) continue;
      if (attribute.name === 'xmlns') return '';
      if (attribute.name.startsWith('xmlns:')) return attribute.name.slice('xmlns:'.length);
    }
  }
  return undefined;
}

function lexicalPrefix(name: string): string {
  const separator = name.indexOf(':');
  return separator < 0 ? '' : name.slice(0, separator);
}

function directChildren(element: XmlElement): XmlElement[] {
  return element.children.filter((child): child is XmlElement => child.type === 'element');
}
