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

const EXTENDED_PROPERTIES_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties';
const EXTENDED_PROPERTIES_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.extended-properties+xml';
const EXTENDED_PROPERTIES_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties';
const FOLLOWING_COMPANY_PROPERTIES = new Set([
  'LinksUpToDate',
  'SharedDoc',
  'HyperlinksChanged',
  'AppVersion',
]);

interface ExtendedPropertiesState {
  readonly relationship: Relationship;
  readonly part: PackagePart;
  readonly xml: LosslessXmlDocument;
  readonly root: XmlElement;
  readonly companies: readonly XmlElement[];
}

export function readPresentationCompany(pkg: OpcPackage): string | undefined {
  const state = readExtendedPropertiesState(pkg);
  if (!state || state.companies.length !== 1) return undefined;
  return readSimpleCompany(state.xml, state.companies[0]!);
}

export function replacePresentationCompany(
  pkg: OpcPackage,
  value: string | undefined,
): void {
  const normalized = normalizePresentationCompany(value);
  const relationships = extendedPropertiesRelationships(pkg);
  if (relationships.length === 0) {
    if (normalized === undefined) return;
    createExtendedProperties(pkg, normalized);
    return;
  }
  if (relationships.length !== 1) {
    throw new PackageError(
      'Presentation has multiple extended-properties relationships',
      '/',
    );
  }

  const state = requireExtendedPropertiesState(pkg, relationships[0]!);
  if (state.companies.length > 1) {
    throw new ModelParseError(
      'Extended properties contain multiple direct companies',
      state.part.uri,
    );
  }
  const company = state.companies[0];
  const current = company ? readSimpleCompany(state.xml, company) : undefined;
  if (company && current === undefined) {
    throw new ModelParseError(
      'Extended properties company is not simple text',
      state.part.uri,
    );
  }
  if (company && normalized !== undefined && current === normalized) return;
  if (!company && normalized === undefined) return;

  if (normalized === undefined) {
    state.xml.removeElement(company!);
  } else if (!company) {
    insertCompany(state.xml, state.root, normalized);
  } else if (company.selfClosing) {
    state.xml.replaceElement(
      company,
      expandSelfClosingCompany(state.xml, company, normalized),
    );
  } else {
    state.xml.replaceText(company, normalized);
  }
  pkg.setPart(state.part.uri, state.xml.serialize(), state.part.contentType);
}

function normalizePresentationCompany(value: unknown): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'string') {
    throw new TypeError('Presentation company must be a string or undefined');
  }
  if (/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(value)) {
    throw new TypeError('Presentation company contains invalid XML characters');
  }
  return value;
}

function extendedPropertiesRelationships(pkg: OpcPackage): readonly Relationship[] {
  return pkg.relationships('/').filter(
    ({ type }) => type === EXTENDED_PROPERTIES_RELATIONSHIP,
  );
}

function readExtendedPropertiesState(
  pkg: OpcPackage,
): ExtendedPropertiesState | undefined {
  const relationships = extendedPropertiesRelationships(pkg);
  if (relationships.length !== 1) return undefined;
  const relationship = relationships[0]!;
  if (relationship.targetMode !== 'Internal' || !relationship.resolvedTarget) return undefined;
  const part = pkg.getPart(relationship.resolvedTarget);
  if (!part || part.contentType !== EXTENDED_PROPERTIES_CONTENT_TYPE) return undefined;
  try {
    return parseExtendedPropertiesState(relationship, part);
  } catch {
    return undefined;
  }
}

function requireExtendedPropertiesState(
  pkg: OpcPackage,
  relationship: Relationship,
): ExtendedPropertiesState {
  if (relationship.targetMode !== 'Internal' || !relationship.resolvedTarget) {
    throw new PackageError('Extended-properties relationship must be internal', '/');
  }
  const part = pkg.getPart(relationship.resolvedTarget);
  if (!part) {
    throw new PackageError(
      'Extended-properties relationship target is missing',
      relationship.resolvedTarget,
    );
  }
  if (part.contentType !== EXTENDED_PROPERTIES_CONTENT_TYPE) {
    throw new PackageError(
      'Extended-properties part has an unsupported content type',
      part.uri,
    );
  }
  try {
    return parseExtendedPropertiesState(relationship, part);
  } catch (error) {
    if (error instanceof ModelParseError) throw error;
    throw new ModelParseError(
      error instanceof Error ? error.message : String(error),
      part.uri,
    );
  }
}

function parseExtendedPropertiesState(
  relationship: Relationship,
  part: PackagePart,
): ExtendedPropertiesState {
  const xml = LosslessXmlDocument.parse(part.bytes);
  if (xml.roots.length !== 1) {
    throw new ModelParseError('Extended-properties part must have one root', part.uri);
  }
  const root = xml.roots[0]!;
  if (
    root.localName !== 'Properties'
    || namespaceUri(root) !== EXTENDED_PROPERTIES_NAMESPACE
  ) {
    throw new ModelParseError('Extended-properties root is invalid', part.uri);
  }
  const companies = directChildren(root).filter(
    (child) => child.localName === 'Company'
      && namespaceUri(child) === EXTENDED_PROPERTIES_NAMESPACE,
  );
  return { relationship, part, xml, root, companies };
}

function readSimpleCompany(
  xml: LosslessXmlDocument,
  company: XmlElement,
): string | undefined {
  if (directChildren(company).length > 0) return undefined;
  if (/<!\[CDATA\[/i.test(xml.original(company))) return undefined;
  return xml.text(company);
}

function createExtendedProperties(pkg: OpcPackage, value: string): void {
  const canonicalUri = '/docProps/app.xml';
  const partUri = pkg.hasPart(canonicalUri)
    ? pkg.allocatePartUri('/docProps', 'app', '.xml')
    : canonicalUri;
  const xml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`
    + `<Properties xmlns="${EXTENDED_PROPERTIES_NAMESPACE}">`
    + `<Company>${escapeXmlText(value)}</Company>`
    + '</Properties>';
  pkg.setPart(partUri, xml, EXTENDED_PROPERTIES_CONTENT_TYPE);
  pkg.addRelationship('/', {
    type: EXTENDED_PROPERTIES_RELATIONSHIP,
    target: partUri.slice(1),
  });
}

function insertCompany(
  xml: LosslessXmlDocument,
  root: XmlElement,
  value: string,
): void {
  const rendered = renderCompany(root, value);
  const following = directChildren(root).find(
    (child) => namespaceUri(child) === EXTENDED_PROPERTIES_NAMESPACE
      && FOLLOWING_COMPANY_PROPERTIES.has(child.localName),
  );
  if (following) {
    xml.replace(following.start, following.start, rendered);
  } else {
    xml.appendChildXml(root, rendered);
  }
}

function renderCompany(root: XmlElement, value: string): string {
  const prefix = lexicalPrefix(root.name);
  const name = prefix === '' ? 'Company' : `${prefix}:Company`;
  return `<${name}>${escapeXmlText(value)}</${name}>`;
}

function expandSelfClosingCompany(
  xml: LosslessXmlDocument,
  company: XmlElement,
  value: string,
): string {
  const original = xml.original(company);
  const marker = original.lastIndexOf('/>');
  if (marker < 0) {
    throw new ModelParseError('Self-closing extended-properties Company is malformed');
  }
  const open = `${original.slice(0, marker).replace(/\s+$/, '')}>`;
  return `${open}${escapeXmlText(value)}</${company.name}>`;
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

function lexicalPrefix(name: string): string {
  const separator = name.indexOf(':');
  return separator < 0 ? '' : name.slice(0, separator);
}

function directChildren(element: XmlElement): XmlElement[] {
  return element.children.filter((child): child is XmlElement => child.type === 'element');
}
