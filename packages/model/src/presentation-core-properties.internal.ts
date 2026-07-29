import {
  escapeXmlAttribute,
  escapeXmlText,
  LosslessXmlDocument,
  type XmlAttribute,
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
  readonly qualifiedType?: CoreQualifiedTypeDescriptor;
}

export interface CoreQualifiedTypeDescriptor {
  readonly attributeNamespace: string;
  readonly attributePreferredPrefix: string;
  readonly valueNamespace: string;
  readonly valuePreferredPrefix: string;
  readonly valueLocalName: string;
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
  const property = state.properties[0]!;
  const value = readSimpleProperty(state.xml, property);
  if (value === undefined) return undefined;
  if (
    descriptor.qualifiedType
    && !qualifiedTypeMatches(property, descriptor.qualifiedType)
  ) {
    return undefined;
  }
  return value;
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
  const qualifiedTypeAttributes = property && descriptor.qualifiedType
    ? attributesByExpandedName(
      property,
      descriptor.qualifiedType.attributeNamespace,
      'type',
    )
    : [];
  if (qualifiedTypeAttributes.length > 1) {
    throw new ModelParseError(
      `Core properties ${descriptor.label} has multiple qualified type attributes`,
      state.part.uri,
    );
  }
  const qualifiedTypeIsCorrect = !descriptor.qualifiedType
    || (property !== undefined && qualifiedTypeMatches(property, descriptor.qualifiedType));
  if (
    property
    && value !== undefined
    && current === value
    && qualifiedTypeIsCorrect
  ) return;
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
      expandSelfClosingProperty(
        state.xml,
        property,
        descriptor,
        value,
        qualifiedTypeAttributes[0],
      ),
    );
  } else if (current !== value) {
    state.xml.replaceText(property, value);
  }
  if (
    property
    && !property.selfClosing
    && descriptor.qualifiedType
    && !qualifiedTypeIsCorrect
  ) {
    ensureQualifiedType(
      state.xml,
      property,
      descriptor.qualifiedType,
      qualifiedTypeAttributes[0],
    );
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
  const rendered = renderNewProperty(
    new Map([['cp', CORE_PROPERTIES_NAMESPACE]]),
    descriptor,
  );
  const xml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`
    + `<cp:coreProperties xmlns:cp="${CORE_PROPERTIES_NAMESPACE}"${renderNamespaceDeclarations(rendered.declarations)}>`
    + `<${rendered.elementName}${rendered.typeAttribute}>${escapeXmlText(value)}</${rendered.elementName}>`
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
  const rendered = renderNewProperty(inScopeNamespaces(root), descriptor);
  return `<${rendered.elementName}${renderNamespaceDeclarations(rendered.declarations)}${rendered.typeAttribute}>`
    + `${escapeXmlText(value)}</${rendered.elementName}>`;
}

function expandSelfClosingProperty(
  xml: LosslessXmlDocument,
  property: XmlElement,
  descriptor: CoreTextPropertyDescriptor,
  value: string,
  qualifiedTypeAttribute: XmlAttribute | undefined,
): string {
  const original = xml.original(property);
  const marker = original.lastIndexOf('/>');
  if (marker < 0) {
    throw new ModelParseError(
      `Self-closing core-properties ${descriptor.label} is malformed`,
    );
  }
  let startTag = original.slice(0, marker).replace(/\s+$/, '');
  if (descriptor.qualifiedType) {
    const bindings = inScopeNamespaces(property);
    const declarations: NamespaceDeclaration[] = [];
    if (qualifiedTypeAttribute) {
      if (!qualifiedTypeValueMatches(property, qualifiedTypeAttribute, descriptor.qualifiedType)) {
        const valuePrefix = selectNamespacePrefix(
          bindings,
          descriptor.qualifiedType.valueNamespace,
          descriptor.qualifiedType.valuePreferredPrefix,
          false,
          declarations,
        );
        const relativeStart = qualifiedTypeAttribute.valueStart - property.start;
        const relativeEnd = qualifiedTypeAttribute.valueEnd - property.start;
        startTag = startTag.slice(0, relativeStart)
          + escapeXmlAttribute(`${valuePrefix}:${descriptor.qualifiedType.valueLocalName}`)
          + startTag.slice(relativeEnd);
      }
    } else {
      const attributePrefix = selectNamespacePrefix(
        bindings,
        descriptor.qualifiedType.attributeNamespace,
        descriptor.qualifiedType.attributePreferredPrefix,
        false,
        declarations,
      );
      const valuePrefix = selectNamespacePrefix(
        bindings,
        descriptor.qualifiedType.valueNamespace,
        descriptor.qualifiedType.valuePreferredPrefix,
        false,
        declarations,
      );
      startTag += `${renderNamespaceDeclarations(declarations)} ${attributePrefix}:type="${escapeXmlAttribute(`${valuePrefix}:${descriptor.qualifiedType.valueLocalName}`)}"`;
      declarations.length = 0;
    }
    startTag += renderNamespaceDeclarations(declarations);
  }
  const open = `${startTag}>`;
  return `${open}${escapeXmlText(value)}</${property.name}>`;
}

interface NamespaceDeclaration {
  readonly prefix: string;
  readonly uri: string;
}

interface RenderedNewProperty {
  readonly elementName: string;
  readonly typeAttribute: string;
  readonly declarations: readonly NamespaceDeclaration[];
}

function renderNewProperty(
  initialBindings: ReadonlyMap<string, string>,
  descriptor: CoreTextPropertyDescriptor,
): RenderedNewProperty {
  const bindings = new Map(initialBindings);
  const declarations: NamespaceDeclaration[] = [];
  const elementPrefix = selectNamespacePrefix(
    bindings,
    descriptor.namespace,
    descriptor.preferredPrefix,
    true,
    declarations,
  );
  const elementName = elementPrefix === ''
    ? descriptor.localName
    : `${elementPrefix}:${descriptor.localName}`;
  if (!descriptor.qualifiedType) {
    return { elementName, typeAttribute: '', declarations };
  }
  const attributePrefix = selectNamespacePrefix(
    bindings,
    descriptor.qualifiedType.attributeNamespace,
    descriptor.qualifiedType.attributePreferredPrefix,
    false,
    declarations,
  );
  const valuePrefix = selectNamespacePrefix(
    bindings,
    descriptor.qualifiedType.valueNamespace,
    descriptor.qualifiedType.valuePreferredPrefix,
    false,
    declarations,
  );
  return {
    elementName,
    typeAttribute: ` ${attributePrefix}:type="${escapeXmlAttribute(`${valuePrefix}:${descriptor.qualifiedType.valueLocalName}`)}"`,
    declarations,
  };
}

function ensureQualifiedType(
  xml: LosslessXmlDocument,
  property: XmlElement,
  descriptor: CoreQualifiedTypeDescriptor,
  attribute: XmlAttribute | undefined,
): void {
  const bindings = inScopeNamespaces(property);
  const declarations: NamespaceDeclaration[] = [];
  if (attribute) {
    const valuePrefix = selectNamespacePrefix(
      bindings,
      descriptor.valueNamespace,
      descriptor.valuePreferredPrefix,
      false,
      declarations,
    );
    xml.replaceAttribute(attribute, `${valuePrefix}:${descriptor.valueLocalName}`);
  } else {
    const attributePrefix = selectNamespacePrefix(
      bindings,
      descriptor.attributeNamespace,
      descriptor.attributePreferredPrefix,
      false,
      declarations,
    );
    const valuePrefix = selectNamespacePrefix(
      bindings,
      descriptor.valueNamespace,
      descriptor.valuePreferredPrefix,
      false,
      declarations,
    );
    const insertionPoint = property.startTagEnd - 1;
    xml.replace(
      insertionPoint,
      insertionPoint,
      `${renderNamespaceDeclarations(declarations)} ${attributePrefix}:type="${escapeXmlAttribute(`${valuePrefix}:${descriptor.valueLocalName}`)}"`,
    );
    return;
  }
  if (declarations.length > 0) {
    const insertionPoint = property.startTagEnd - 1;
    xml.replace(
      insertionPoint,
      insertionPoint,
      renderNamespaceDeclarations(declarations),
    );
  }
}

function qualifiedTypeMatches(
  property: XmlElement,
  descriptor: CoreQualifiedTypeDescriptor,
): boolean {
  const attributes = attributesByExpandedName(
    property,
    descriptor.attributeNamespace,
    'type',
  );
  return attributes.length === 1
    && qualifiedTypeValueMatches(property, attributes[0]!, descriptor);
}

function qualifiedTypeValueMatches(
  property: XmlElement,
  attribute: XmlAttribute,
  descriptor: CoreQualifiedTypeDescriptor,
): boolean {
  const separator = attribute.value.indexOf(':');
  if (
    separator <= 0
    || separator !== attribute.value.lastIndexOf(':')
    || separator === attribute.value.length - 1
  ) return false;
  const prefix = attribute.value.slice(0, separator);
  const localName = attribute.value.slice(separator + 1);
  return localName === descriptor.valueLocalName
    && namespaceUriForPrefix(property, prefix) === descriptor.valueNamespace;
}

function attributesByExpandedName(
  element: XmlElement,
  namespace: string,
  localName: string,
): readonly XmlAttribute[] {
  return element.attributes.filter(
    (attribute) => attribute.localName === localName
      && attributeNamespaceUri(element, attribute) === namespace,
  );
}

function attributeNamespaceUri(
  element: XmlElement,
  attribute: XmlAttribute,
): string | undefined {
  const prefix = lexicalPrefix(attribute.name);
  return prefix === '' ? undefined : namespaceUriForPrefix(element, prefix);
}

function namespaceUri(element: XmlElement): string | undefined {
  return namespaceUriForPrefix(element, lexicalPrefix(element.name));
}

function namespaceUriForPrefix(element: XmlElement, prefix: string): string | undefined {
  return inScopeNamespaces(element).get(prefix);
}

function inScopeNamespaces(element: XmlElement): Map<string, string> {
  const bindings = new Map<string, string>();
  for (let scope: XmlElement | undefined = element; scope; scope = scope.parent) {
    for (const attribute of scope.attributes) {
      const prefix = attribute.name === 'xmlns'
        ? ''
        : attribute.name.startsWith('xmlns:')
          ? attribute.name.slice('xmlns:'.length)
          : undefined;
      if (prefix !== undefined && !bindings.has(prefix)) {
        bindings.set(prefix, attribute.value);
      }
    }
  }
  return bindings;
}

function selectNamespacePrefix(
  bindings: Map<string, string>,
  uri: string,
  preferredPrefix: string,
  allowDefault: boolean,
  declarations: NamespaceDeclaration[],
): string {
  for (const [prefix, namespace] of bindings) {
    if (namespace === uri && (allowDefault || prefix !== '')) return prefix;
  }
  let prefix = preferredPrefix;
  for (let suffix = 1; bindings.has(prefix); suffix += 1) {
    prefix = `${preferredPrefix}${suffix}`;
  }
  bindings.set(prefix, uri);
  declarations.push({ prefix, uri });
  return prefix;
}

function renderNamespaceDeclarations(
  declarations: readonly NamespaceDeclaration[],
): string {
  return declarations
    .map(({ prefix, uri }) => ` xmlns:${prefix}="${escapeXmlAttribute(uri)}"`)
    .join('');
}

function lexicalPrefix(name: string): string {
  const separator = name.indexOf(':');
  return separator < 0 ? '' : name.slice(0, separator);
}

function directChildren(element: XmlElement): XmlElement[] {
  return element.children.filter((child): child is XmlElement => child.type === 'element');
}
