import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import type { OpcPackage } from '@pptx/opc';
import { ModelParseError } from './errors.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const CONTENT_TYPE_NAMESPACE =
  'application/vnd.openxmlformats-officedocument.presentationml.';
const SLIDE_MASTER_RELATIONSHIP = `${RELATIONSHIP_NAMESPACE}slideMaster`;
const SLIDE_LAYOUT_RELATIONSHIP = `${RELATIONSHIP_NAMESPACE}slideLayout`;
const SLIDE_MASTER_CONTENT_TYPE = `${CONTENT_TYPE_NAMESPACE}slideMaster+xml`;
const SLIDE_LAYOUT_CONTENT_TYPE = `${CONTENT_TYPE_NAMESPACE}slideLayout+xml`;

interface AttachedSlideLayout {
  readonly partUri: string;
  readonly name?: string;
}

export function resolveSlideLayoutPartUri(
  pkg: OpcPackage,
  presentationPartUri: string,
  masterName: string | undefined,
  inheritedLayoutPartUri: string | undefined,
): string | undefined {
  const layouts = attachedSlideLayouts(pkg, presentationPartUri);
  if (masterName !== undefined) {
    const matches = layouts.filter(({ name }) => name === masterName);
    if (matches.length === 0) {
      throw new RangeError(`Slide master ${masterName} was not found`);
    }
    if (matches.length !== 1) {
      throw new ModelParseError(
        `Slide master ${masterName} is ambiguous`,
        presentationPartUri,
      );
    }
    return matches[0]!.partUri;
  }

  if (inheritedLayoutPartUri !== undefined) {
    const inherited = layouts.filter(
      ({ partUri }) => partUri === inheritedLayoutPartUri,
    );
    if (inherited.length === 1) return inherited[0]!.partUri;
    if (inherited.length > 1) {
      throw new ModelParseError(
        'Inherited slide layout is ambiguous',
        presentationPartUri,
      );
    }
  }
  return layouts[0]?.partUri;
}

function attachedSlideLayouts(
  pkg: OpcPackage,
  presentationPartUri: string,
): AttachedSlideLayout[] {
  const layouts: AttachedSlideLayout[] = [];
  const masters = pkg.relationships(presentationPartUri).filter((relationship) => {
    if (
      relationship.type !== SLIDE_MASTER_RELATIONSHIP
      || relationship.targetMode !== 'Internal'
      || !relationship.resolvedTarget
    ) return false;
    return pkg.getPart(relationship.resolvedTarget)?.contentType === SLIDE_MASTER_CONTENT_TYPE;
  });

  for (const master of masters) {
    for (const relationship of pkg.relationships(master.resolvedTarget!)) {
      if (
        relationship.type !== SLIDE_LAYOUT_RELATIONSHIP
        || relationship.targetMode !== 'Internal'
        || !relationship.resolvedTarget
      ) continue;
      const part = pkg.getPart(relationship.resolvedTarget);
      if (part?.contentType !== SLIDE_LAYOUT_CONTENT_TYPE) continue;
      const parsed = directLayoutName(part.bytes, relationship.resolvedTarget);
      if (parsed.valid) {
        layouts.push({
          partUri: relationship.resolvedTarget,
          ...(parsed.name !== undefined ? { name: parsed.name } : {}),
        });
      }
    }
  }
  return layouts;
}

function directLayoutName(
  bytes: Uint8Array,
  partUri: string,
): { readonly valid: boolean; readonly name?: string } {
  let xml: LosslessXmlDocument;
  try {
    xml = LosslessXmlDocument.parse(bytes);
  } catch {
    throw new ModelParseError('Invalid slide layout XML', partUri);
  }
  if (xml.roots.length !== 1) return { valid: false };
  const root = xml.roots[0]!;
  if (
    root.localName !== 'sldLayout'
    || namespaceUri(root) !== PRESENTATION_NAMESPACE
  ) return { valid: false };
  const commonSlideData = directChildren(root).filter(
    (child) => child.localName === 'cSld'
      && namespaceUri(child) === PRESENTATION_NAMESPACE,
  );
  if (commonSlideData.length !== 1) return { valid: false };
  const names = commonSlideData[0]!.attributes.filter(({ name }) => name === 'name');
  if (names.length > 1) return { valid: false };
  return {
    valid: true,
    ...(names[0] ? { name: names[0].value } : {}),
  };
}

function directChildren(element: XmlElement): XmlElement[] {
  return element.children.filter(
    (child): child is XmlElement => child.type === 'element',
  );
}

function namespaceUri(element: XmlElement): string | undefined {
  return namespaceUriForPrefix(element, lexicalPrefix(element.name));
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
