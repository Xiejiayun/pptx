import { type LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import { type OpcPackage, type Relationship } from '@pptx/opc';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const MEDIA_REL = 'http://schemas.microsoft.com/office/2007/relationships/media';
const IMAGE_REL = `${REL}image`;
const PLAYBACK_EXTENSION_URI = '{C13D3E4A-5148-4B6D-A7E7-505054582D4F}';

export type MediaStateKind = 'audio' | 'video';

export interface MediaStatePlaybackSettings {
  readonly play?: 'click' | 'auto';
  readonly loop?: boolean;
  readonly hideWhenStopped?: boolean;
  readonly volume?: number;
}

export interface MediaState {
  readonly kind: MediaStateKind;
  readonly shapeId: number;
  readonly slidePartUri: string;
  readonly name: string;
  readonly altText?: string;
  readonly mediaPartUri?: string;
  readonly externalUrl?: string;
  readonly posterPartUri?: string;
  readonly settings: MediaStatePlaybackSettings;
}

export function readMediaState(
  pkg: OpcPackage,
  slidePartUri: string,
  xml: LosslessXmlDocument,
  picture: XmlElement,
): Readonly<MediaState> | undefined {
  if (picture.localName !== 'pic') return undefined;
  const properties = xml.descendants(picture, 'cNvPr')[0];
  if (!properties) return undefined;
  const shapeIdText = xml.attribute(properties, 'id')?.value;
  if (!shapeIdText || !/^\d+$/.test(shapeIdText)) return undefined;
  const shapeId = Number(shapeIdText);
  if (!Number.isSafeInteger(shapeId)) return undefined;

  const marker = xml.descendants(picture)
    .find(({ localName }) => localName === 'audioFile' || localName === 'videoFile');
  const markerKind = marker?.localName === 'audioFile'
    ? 'audio'
    : marker?.localName === 'videoFile'
      ? 'video'
      : undefined;
  const markerRelationship = relationshipForAttribute(
    pkg,
    slidePartUri,
    xml,
    marker,
    'r:link',
  );
  const mediaExtension = xml.descendants(picture, 'media')[0];
  const extensionRelationship = relationshipForAttribute(
    pkg,
    slidePartUri,
    xml,
    mediaExtension,
    'r:embed',
  );
  const referencedRelationships = uniqueRelationships([
    markerRelationship,
    extensionRelationship,
  ]);
  const standardRelationship = referencedRelationships.find(standardMediaRelationshipKind);
  const kind = standardRelationship
    ? standardMediaRelationshipKind(standardRelationship)
    : kindFromTargetContentType(pkg, referencedRelationships) ?? markerKind;
  if (!kind) return undefined;

  const primary = choosePrimaryRelationship(pkg, kind, referencedRelationships, standardRelationship);
  const blip = xml.descendants(picture, 'blip')[0];
  const poster = relationshipForAttribute(pkg, slidePartUri, xml, blip, 'r:embed');
  const altText = xml.attribute(properties, 'descr')?.value;
  const settings = readPlaybackSettings(xml, picture);
  const state: MediaState = {
    kind,
    shapeId,
    slidePartUri,
    name: xml.attribute(properties, 'name')?.value ?? `Media ${shapeId}`,
    ...(altText !== undefined ? { altText } : {}),
    ...(primary?.targetMode === 'Internal' && primary.resolvedTarget
      ? { mediaPartUri: primary.resolvedTarget }
      : {}),
    ...(primary?.targetMode === 'External' ? { externalUrl: primary.target } : {}),
    ...(poster?.type === IMAGE_REL && poster.targetMode === 'Internal' && poster.resolvedTarget
      ? { posterPartUri: poster.resolvedTarget }
      : {}),
    settings,
  };
  return Object.freeze(state);
}

function relationshipForAttribute(
  pkg: OpcPackage,
  slidePartUri: string,
  xml: LosslessXmlDocument,
  element: XmlElement | undefined,
  attributeName: string,
): Relationship | undefined {
  const id = element ? xml.attribute(element, attributeName)?.value : undefined;
  return id ? pkg.relationships(slidePartUri).find((relationship) => relationship.id === id) : undefined;
}

function uniqueRelationships(
  values: readonly (Relationship | undefined)[],
): readonly Relationship[] {
  const seen = new Set<string>();
  return values.filter((value): value is Relationship => {
    if (!value || seen.has(value.id)) return false;
    seen.add(value.id);
    return true;
  });
}

function standardMediaRelationshipKind(
  relationship: Relationship,
): MediaStateKind | undefined {
  if (relationship.type === `${REL}audio`) return 'audio';
  if (relationship.type === `${REL}video`) return 'video';
  return undefined;
}

function kindFromTargetContentType(
  pkg: OpcPackage,
  relationships: readonly Relationship[],
): MediaStateKind | undefined {
  for (const relationship of relationships) {
    if (relationship.targetMode !== 'Internal' || !relationship.resolvedTarget) continue;
    const contentType = pkg.getPart(relationship.resolvedTarget)?.contentType;
    if (contentType?.startsWith('audio/')) return 'audio';
    if (contentType?.startsWith('video/')) return 'video';
  }
  return undefined;
}

function choosePrimaryRelationship(
  pkg: OpcPackage,
  kind: MediaStateKind,
  relationships: readonly Relationship[],
  standardRelationship: Relationship | undefined,
): Relationship | undefined {
  if (standardRelationship) return standardRelationship;
  const typed = relationships.find((relationship) => {
    if (relationship.targetMode !== 'Internal' || !relationship.resolvedTarget) return false;
    return pkg.getPart(relationship.resolvedTarget)?.contentType.startsWith(`${kind}/`) ?? false;
  });
  if (typed) return typed;
  return relationships.find(({ type }) => type === MEDIA_REL) ?? relationships[0];
}

function readPlaybackSettings(
  xml: LosslessXmlDocument,
  picture: XmlElement,
): Readonly<MediaStatePlaybackSettings> {
  const extension = xml.descendants(picture, 'ext').find(
    (candidate) => xml.attribute(candidate, 'uri')?.value === PLAYBACK_EXTENSION_URI,
  );
  const playback = extension
    ? xml.descendants(extension, 'playback')[0]
    : undefined;
  if (!playback) return Object.freeze({});
  const playValue = xml.attribute(playback, 'play')?.value;
  const loopValue = booleanToken(xml.attribute(playback, 'loop')?.value);
  const hideValue = booleanToken(xml.attribute(playback, 'hideWhenStopped')?.value);
  const volumeValue = xml.attribute(playback, 'volume')?.value;
  const volume = volumeValue && /^\d+$/.test(volumeValue)
    ? Number(volumeValue) / 100_000
    : undefined;
  return Object.freeze({
    ...(playValue === 'click' || playValue === 'auto' ? { play: playValue } : {}),
    ...(loopValue !== undefined ? { loop: loopValue } : {}),
    ...(hideValue !== undefined ? { hideWhenStopped: hideValue } : {}),
    ...(volume !== undefined && Number.isFinite(volume) && volume >= 0 && volume <= 1
      ? { volume }
      : {}),
  });
}

function booleanToken(value: string | undefined): boolean | undefined {
  if (value === '1' || value === 'true' || value === 'on') return true;
  if (value === '0' || value === 'false' || value === 'off') return false;
  return undefined;
}
