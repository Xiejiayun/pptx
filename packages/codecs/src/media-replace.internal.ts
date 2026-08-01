import {
  escapeXmlAttribute,
  LosslessXmlDocument,
  type XmlElement,
} from '@pptx/lossless-xml';
import {
  partUriExtension,
  relativeRelationshipTarget,
  type OpcPackage,
  type Relationship,
} from '@pptx/opc';
import { normalizeMediaCreateRequest, type NormalizedMediaSourceReference } from './media-create.internal.js';
import {
  resolveMediaSourceInput,
  type ResolvedEmbeddedMedia,
  type ResolvedExternalMedia,
} from './media-source.internal.js';
import { readMediaState } from './media-state.internal.js';
import type { MediaKind, MediaSource, ReplaceMediaSourceOptions } from './media.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const MEDIA_REL = 'http://schemas.microsoft.com/office/2007/relationships/media';
const OFFICE_MEDIA_EXTENSION_URI = '{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}';
const OPTION_KEYS = new Set(['contentType', 'fileName', 'transcode']);

export interface NormalizedMediaReplaceRequest {
  readonly kind: MediaKind;
  readonly source: NormalizedMediaSourceReference;
  readonly contentType?: string;
  readonly fileName?: string;
  readonly transcode?: NonNullable<ReplaceMediaSourceOptions['transcode']>;
}

export function normalizeMediaReplaceRequest(
  kind: MediaKind,
  source: MediaSource,
  options: ReplaceMediaSourceOptions = {},
): Readonly<NormalizedMediaReplaceRequest> {
  const values = readOptions(options);
  const creation = normalizeMediaCreateRequest(kind, source, {
    ...(values.contentType === undefined ? {} : { contentType: values.contentType }),
    ...(values.fileName === undefined ? {} : { fileName: values.fileName }),
    ...(values.transcode === undefined ? {} : { transcode: values.transcode }),
  });
  return Object.freeze({
    kind: creation.kind,
    source: creation.source,
    ...(creation.contentType === undefined ? {} : { contentType: creation.contentType }),
    ...(creation.fileName === undefined ? {} : { fileName: creation.fileName }),
    ...(creation.transcode === undefined ? {} : { transcode: creation.transcode }),
  });
}

export async function resolveMediaReplacementSource(
  request: Readonly<NormalizedMediaReplaceRequest>,
): Promise<ResolvedEmbeddedMedia | ResolvedExternalMedia> {
  return resolveMediaSourceInput(request);
}

export async function findMatchingMediaPart(
  pkg: OpcPackage,
  value: Readonly<ResolvedEmbeddedMedia>,
): Promise<string | undefined> {
  const expected = await hash(value.bytes);
  for (const part of pkg.parts) {
    if (
      part.uri.startsWith('/ppt/media/')
      && part.contentType === value.contentType
      && await hash(part.bytes) === expected
    ) return part.uri;
  }
  return undefined;
}

export function replaceResolvedMediaSource(
  pkg: OpcPackage,
  slidePartUri: string,
  shapeId: number,
  kind: MediaKind,
  resolved: Readonly<ResolvedEmbeddedMedia | ResolvedExternalMedia>,
  matchingPartUri?: string,
): void {
  pkg.transaction(() => {
    const slidePart = pkg.requirePart(slidePartUri);
    const xml = LosslessXmlDocument.parse(slidePart.bytes);
    const picture = requirePicture(xml, shapeId, slidePartUri);
    const currentState = readMediaState(pkg, slidePartUri, xml, picture);
    if (!currentState || currentState.kind !== kind) {
      throw new Error(`Media shape ${shapeId} is not ${kind} on ${slidePartUri}`);
    }
    const roles = readPrimaryRoles(pkg, slidePartUri, xml, picture);
    const oldIds = new Set([roles.kindId, roles.mediaId].filter((id): id is string => Boolean(id)));
    const oldTargets = new Set(
      [roles.kindRelationship?.resolvedTarget, roles.mediaRelationship?.resolvedTarget]
        .filter((target): target is string => Boolean(target)),
    );

    let targetUri: string | undefined;
    if (resolved.type === 'embedded') {
      const currentTarget = roles.kindRelationship?.targetMode === 'Internal'
        ? roles.kindRelationship.resolvedTarget
        : roles.mediaRelationship?.resolvedTarget;
      const canMutate = currentTarget !== undefined
        && matchingPartUri === undefined
        && partUriExtension(currentTarget) === resolved.extension
        && targetOwnedByPicture(pkg, slidePartUri, xml, currentTarget, roles);
      targetUri = matchingPartUri ?? (canMutate ? currentTarget : undefined)
        ?? pkg.allocatePartUri('/ppt/media', 'media', resolved.extension);
      if (!pkg.hasPart(targetUri) || targetUri === currentTarget) {
        pkg.setPart(targetUri, resolved.bytes, resolved.contentType);
      }
    }

    const kindRelationship = upsertRoleRelationship(pkg, slidePartUri, xml, roles.kindRelationship, {
      type: `${REL}${kind}`,
      target: resolved.type === 'external'
        ? resolved.url
        : relativeRelationshipTarget(slidePartUri, targetUri!),
      targetMode: resolved.type === 'external' ? 'External' : 'Internal',
    });
    const mediaRelationship = resolved.type === 'embedded'
      ? upsertRoleRelationship(
          pkg,
          slidePartUri,
          xml,
          roles.mediaRelationship?.id === kindRelationship.id ? undefined : roles.mediaRelationship,
          {
            type: MEDIA_REL,
            target: relativeRelationshipTarget(slidePartUri, targetUri!),
            targetMode: 'Internal',
          },
        )
      : undefined;

    replaceMarker(xml, picture, roles.marker, kind, kindRelationship.id);
    replaceOfficeMediaExtension(xml, picture, roles.mediaExtension, mediaRelationship?.id);
    const updatedXml = xml.serialize();
    pkg.setPart(slidePartUri, updatedXml, slidePart.contentType);

    const reparsed = LosslessXmlDocument.parse(updatedXml);
    for (const id of oldIds) {
      if (id !== kindRelationship.id && id !== mediaRelationship?.id && relationshipReferenceCount(reparsed, id) === 0) {
        pkg.removeRelationship(slidePartUri, id);
      }
    }
    for (const target of oldTargets) {
      if (target === targetUri || !target.startsWith('/ppt/media/')) continue;
      const incoming = pkg.graph.find(({ uri }) => uri === target)?.incoming ?? [];
      if (incoming.length === 0) pkg.deletePart(target);
    }
  });
}

interface PrimaryRoles {
  readonly marker?: XmlElement;
  readonly kindId?: string;
  readonly kindRelationship?: Relationship;
  readonly mediaExtension?: XmlElement;
  readonly mediaId?: string;
  readonly mediaRelationship?: Relationship;
}

function readPrimaryRoles(
  pkg: OpcPackage,
  slidePartUri: string,
  xml: LosslessXmlDocument,
  picture: XmlElement,
): PrimaryRoles {
  const marker = xml.descendants(picture).find(
    ({ localName }) => localName === 'audioFile' || localName === 'videoFile',
  );
  const mediaExtension = xml.descendants(picture, 'media')[0];
  const kindId = marker ? xml.attribute(marker, 'r:link')?.value : undefined;
  const mediaId = mediaExtension ? xml.attribute(mediaExtension, 'r:embed')?.value : undefined;
  const relationships = pkg.relationships(slidePartUri);
  const kindRelationship = kindId
    ? relationships.find(({ id }) => id === kindId)
    : undefined;
  const mediaRelationship = mediaId
    ? relationships.find(({ id }) => id === mediaId)
    : undefined;
  return {
    ...(marker ? { marker } : {}),
    ...(kindId ? { kindId } : {}),
    ...(kindRelationship ? { kindRelationship } : {}),
    ...(mediaExtension ? { mediaExtension } : {}),
    ...(mediaId ? { mediaId } : {}),
    ...(mediaRelationship ? { mediaRelationship } : {}),
  };
}

function upsertRoleRelationship(
  pkg: OpcPackage,
  slidePartUri: string,
  xml: LosslessXmlDocument,
  current: Relationship | undefined,
  value: { type: string; target: string; targetMode: 'Internal' | 'External' },
): Relationship {
  if (current && relationshipReferenceCount(xml, current.id) === 1) {
    return pkg.updateRelationship(slidePartUri, current.id, value);
  }
  return pkg.addRelationship(slidePartUri, value);
}

function targetOwnedByPicture(
  pkg: OpcPackage,
  slidePartUri: string,
  xml: LosslessXmlDocument,
  target: string,
  roles: PrimaryRoles,
): boolean {
  const roleIds = new Set([roles.kindId, roles.mediaId].filter((id): id is string => Boolean(id)));
  if (roleIds.size === 0) return false;
  const incoming = pkg.graph.find(({ uri }) => uri === target)?.incoming ?? [];
  return incoming.length > 0 && incoming.every(
    ({ sourceUri, relationship }) => sourceUri === slidePartUri
      && roleIds.has(relationship.id)
      && relationshipReferenceCount(xml, relationship.id) === 1,
  );
}

function replaceMarker(
  xml: LosslessXmlDocument,
  picture: XmlElement,
  current: XmlElement | undefined,
  kind: MediaKind,
  relationshipId: string,
): void {
  const rendered = `<a:${kind}File r:link="${escapeXmlAttribute(relationshipId)}"/>`;
  if (current) xml.replace(current.start, current.end, rendered);
  else {
    const properties = xml.descendants(picture, 'nvPr')[0];
    if (!properties) throw new Error('Media picture has no application properties');
    xml.replace(properties.startTagEnd, properties.startTagEnd, rendered);
  }
}

function replaceOfficeMediaExtension(
  xml: LosslessXmlDocument,
  picture: XmlElement,
  currentMedia: XmlElement | undefined,
  relationshipId: string | undefined,
): void {
  const currentExtension = currentMedia?.parent?.localName === 'ext' ? currentMedia.parent : undefined;
  if (relationshipId === undefined) {
    if (currentMedia && currentExtension) {
      const siblings = currentExtension.children.filter(
        (child): child is XmlElement => child.type === 'element' && child !== currentMedia,
      );
      if (siblings.length === 0) xml.removeElement(currentExtension);
      else xml.removeElement(currentMedia);
    }
    return;
  }
  const mediaXml = '<p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" '
    + `r:embed="${escapeXmlAttribute(relationshipId)}"/>`;
  if (currentMedia && currentExtension) {
    const uri = xml.attribute(currentExtension, 'uri');
    if (uri) xml.replaceAttribute(uri, OFFICE_MEDIA_EXTENSION_URI);
    else insertAttribute(xml, currentExtension, 'uri', OFFICE_MEDIA_EXTENSION_URI);
    xml.replace(currentMedia.start, currentMedia.end, mediaXml);
    return;
  }
  const rendered = `<p:ext uri="${OFFICE_MEDIA_EXTENSION_URI}">${mediaXml}</p:ext>`;
  const properties = xml.descendants(picture, 'nvPr')[0];
  if (!properties) throw new Error('Media picture has no application properties');
  const extensionList = directChildren(properties, 'extLst')[0];
  const ownedExtension = extensionList
    ? directChildren(extensionList, 'ext').find(
        (extension) => xml.attribute(extension, 'uri')?.value === OFFICE_MEDIA_EXTENSION_URI,
      )
    : undefined;
  if (ownedExtension) xml.appendChildXml(ownedExtension, mediaXml);
  else if (extensionList) xml.appendChildXml(extensionList, rendered);
  else xml.appendChildXml(properties, `<p:extLst>${rendered}</p:extLst>`);
}

function insertAttribute(
  xml: LosslessXmlDocument,
  element: XmlElement,
  name: string,
  value: string,
): void {
  const position = element.selfClosing
    ? xml.source.lastIndexOf('/', element.startTagEnd - 1)
    : element.startTagEnd - 1;
  if (position <= element.start) throw new Error(`Media ${element.localName} start tag is invalid`);
  xml.replace(position, position, ` ${name}="${escapeXmlAttribute(value)}"`);
}

function requirePicture(
  xml: LosslessXmlDocument,
  shapeId: number,
  slidePartUri: string,
): XmlElement {
  const picture = xml.elements('pic').find((candidate) => {
    const properties = xml.descendants(candidate, 'cNvPr')[0];
    return Number(properties ? xml.attribute(properties, 'id')?.value : -1) === shapeId;
  });
  if (!picture) throw new Error(`Media shape ${shapeId} was not found on ${slidePartUri}`);
  return picture;
}

function relationshipReferenceCount(xml: LosslessXmlDocument, id: string): number {
  return xml.elements().flatMap(({ attributes }) => attributes)
    .filter(({ name, value }) => name.startsWith('r:') && value === id).length;
}

function directChildren(element: XmlElement, localName: string): XmlElement[] {
  return element.children.filter(
    (child): child is XmlElement => child.type === 'element' && child.localName === localName,
  );
}

function readOptions(value: unknown): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError('Media replacement options must be an object');
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError('Media replacement options must be an ordinary object');
  }
  const result = Object.create(null) as Record<string, unknown>;
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !OPTION_KEYS.has(key)) {
      throw new TypeError(`Media replacement options contain unsupported property ${String(key)}`);
    }
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`Media replacement option ${key} must be a data property`);
    }
    result[key] = descriptor.value;
  }
  return result;
}

async function hash(bytes: Uint8Array): Promise<string> {
  if (!globalThis.crypto?.subtle) throw new Error('Media hashing requires the Web Crypto API');
  const digest = await globalThis.crypto.subtle.digest('SHA-256', new Uint8Array(bytes).buffer);
  return [...new Uint8Array(digest)].map((byte) => byte.toString(16).padStart(2, '0')).join('');
}
