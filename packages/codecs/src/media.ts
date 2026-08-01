import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import { relativeRelationshipTarget, type OpcPackage } from '@pptx/opc';
import {
  finalizeMediaCreationDefinition,
  normalizeMediaCreateRequest,
  renderMediaPictureXml,
} from './media-create.internal.js';
import {
  mediaPlaybackSettingsEqual,
  normalizeMediaAltText,
  normalizeMediaName,
  normalizeMediaPlaybackSettings,
  replaceMediaMetadataAttribute,
  replaceMediaPlaybackExtension,
} from './media-edit.internal.js';
import { resolveMediaCreationInputs } from './media-source.internal.js';
import { readMediaState } from './media-state.internal.js';
import type { CodecDiagnostic } from './registry.js';

export {
  readMediaState,
  type MediaState,
  type MediaStateKind,
  type MediaStatePlaybackSettings,
} from './media-state.internal.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const MEDIA_REL = 'http://schemas.microsoft.com/office/2007/relationships/media';

export type MediaKind = 'audio' | 'video';
export type MediaByteChunk = number | Uint8Array | ArrayBuffer | ArrayBufferView;
export type MediaByteStream = ReadableStream<MediaByteChunk> | AsyncIterable<MediaByteChunk>;
export type MediaSource = string | Uint8Array | ArrayBuffer | Blob | MediaByteStream;

export interface MediaPlaybackSettings {
  readonly play?: 'click' | 'auto';
  readonly loop?: boolean;
  readonly hideWhenStopped?: boolean;
  readonly volume?: number;
}

export interface AddMediaOptions extends MediaPlaybackSettings {
  readonly name?: string;
  readonly altText?: string;
  readonly contentType?: string;
  readonly fileName?: string;
  readonly poster?: MediaSource;
  readonly posterContentType?: string;
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
  readonly transcode?: (
    bytes: Uint8Array,
    contentType: string,
    kind: MediaKind,
  ) => Promise<{ bytes: Uint8Array; contentType: string; extension?: string }>;
}

export interface MediaDescriptor {
  readonly kind: MediaKind;
  readonly shapeId: number;
  readonly slidePartUri: string;
  readonly mediaPartUri?: string;
  readonly externalUrl?: string;
  readonly posterPartUri?: string;
  readonly settings: MediaPlaybackSettings;
}

/** @deprecated Use the runtime MediaModel from @pptx/model or @pptx/sdk. */
export type MediaModel = MediaDescriptor;

export class MediaCodec {
  readonly id = 'builtin.media';
  readonly priority = 100;
  readonly ownership = {
    elements: ['a:audioFile', 'a:videoFile', 'p14:media'],
    relationshipTypes: [`${REL}audio`, `${REL}video`, MEDIA_REL],
  } as const;

  constructor(readonly pkg: OpcPackage) {}

  async addAudio(slidePartUri: string, source: MediaSource, options: AddMediaOptions = {}): Promise<MediaDescriptor> {
    return this.add('audio', slidePartUri, source, options);
  }

  async addVideo(slidePartUri: string, source: MediaSource, options: AddMediaOptions = {}): Promise<MediaDescriptor> {
    return this.add('video', slidePartUri, source, options);
  }

  list(slidePartUri: string): readonly MediaDescriptor[] {
    const xml = LosslessXmlDocument.parse(this.pkg.requirePart(slidePartUri).bytes);
    return xml.elements('pic')
      .map((picture) => readMediaState(this.pkg, slidePartUri, xml, picture))
      .filter((state): state is NonNullable<typeof state> => state !== undefined);
  }

  setName(slidePartUri: string, shapeId: number, value: string): void {
    const normalized = normalizeMediaName(value);
    this.editPicture(slidePartUri, shapeId, (xml, picture) =>
      replaceMediaMetadataAttribute(xml, picture, 'name', normalized));
  }

  setAltText(slidePartUri: string, shapeId: number, value: string | undefined): void {
    const normalized = normalizeMediaAltText(value);
    this.editPicture(slidePartUri, shapeId, (xml, picture) =>
      replaceMediaMetadataAttribute(xml, picture, 'descr', normalized));
  }

  setSettings(
    slidePartUri: string,
    shapeId: number,
    value: MediaPlaybackSettings | undefined,
  ): void {
    const normalized = normalizeMediaPlaybackSettings(value);
    this.editPicture(slidePartUri, shapeId, (xml, picture) => {
      const state = readMediaState(this.pkg, slidePartUri, xml, picture)!;
      if (normalized && mediaPlaybackSettingsEqual(state.settings, normalized)) return false;
      return replaceMediaPlaybackExtension(xml, picture, normalized);
    });
  }

  delete(slidePartUri: string, shapeId: number): void {
    this.pkg.transaction(() => {
      const part = this.pkg.requirePart(slidePartUri);
      const xml = LosslessXmlDocument.parse(part.bytes);
      const picture = xml.elements('pic').find((candidate) => {
        const properties = xml.descendants(candidate, 'cNvPr')[0];
        return Number(properties ? xml.attribute(properties, 'id')?.value : -1) === shapeId;
      });
      if (!picture) throw new Error(`Media shape ${shapeId} was not found on ${slidePartUri}`);
      const ids = new Set(
        xml
          .descendants(picture)
          .flatMap((element) => element.attributes)
          .filter(({ name }) => name === 'r:embed' || name === 'r:link')
          .map(({ value }) => value),
      );
      const targets = this.pkg
        .relationships(slidePartUri)
        .filter(({ id }) => ids.has(id))
        .map(({ resolvedTarget }) => resolvedTarget)
        .filter((target): target is string => Boolean(target));
      xml.removeElement(picture);
      this.pkg.setPart(slidePartUri, xml.serialize(), part.contentType);
      for (const id of ids) this.pkg.removeRelationship(slidePartUri, id);
      for (const target of new Set(targets)) {
        const incoming = this.pkg.graph.find(({ uri }) => uri === target)?.incoming ?? [];
        if (incoming.length === 0 && target.startsWith('/ppt/media/')) this.pkg.deletePart(target);
      }
    });
  }

  diagnostics(model: MediaDescriptor, profile: string): CodecDiagnostic[] {
    const diagnostics: CodecDiagnostic[] = [];
    if (model.externalUrl) {
      diagnostics.push({
        severity: 'warning',
        code: 'MEDIA_EXTERNAL_NOT_PORTABLE',
        message: `External ${model.kind} may be unavailable when the presentation moves`,
        partUri: model.slidePartUri,
      });
    }
    if (model.settings.play === 'auto' && profile === 'google-slides-import') {
      diagnostics.push({
        severity: 'warning',
        code: 'MEDIA_AUTOPLAY_MAY_DEGRADE',
        message: 'Google Slides import may change autoplay behavior',
        partUri: model.slidePartUri,
      });
    }
    if (
      model.settings.play === 'auto' ||
      model.settings.loop ||
      model.settings.hideWhenStopped ||
      (model.settings.volume !== undefined && model.settings.volume !== 1)
    ) {
      const xml = LosslessXmlDocument.parse(this.pkg.requirePart(model.slidePartUri).bytes);
      const hasNativeTiming = xml.elements('spTgt').some(
        (target) =>
          Number(xml.attribute(target, 'spid')?.value) === model.shapeId && Boolean(ancestor(target, 'cMediaNode')),
      );
      if (!hasNativeTiming) {
        diagnostics.push({
          severity: 'info',
          code: 'MEDIA_PLAYBACK_TIMING_EXTENSION',
          message: 'Playback preferences are preserved and require the timing codec for native client behavior',
          partUri: model.slidePartUri,
        });
      }
    }
    if (model.mediaPartUri) {
      const contentType = this.pkg.getPart(model.mediaPartUri)?.contentType;
      if (profile === 'powerpoint-2010' && (contentType === 'video/webm' || contentType === 'audio/ogg')) {
        diagnostics.push({
          severity: 'warning',
          code: 'MEDIA_CODEC_MAY_BE_UNSUPPORTED',
          message: `${contentType} is not a reliable PowerPoint 2010 media format`,
          partUri: model.mediaPartUri,
        });
      }
    }
    return diagnostics;
  }

  private async add(
    kind: MediaKind,
    slidePartUri: string,
    source: MediaSource,
    options: AddMediaOptions,
  ): Promise<MediaDescriptor> {
    const request = normalizeMediaCreateRequest(kind, source, options);
    const resolved = await resolveMediaCreationInputs(request);
    const existingMediaPartUri = resolved.media.type === 'embedded'
      ? await this.findByHash(resolved.media.bytes, resolved.media.contentType)
      : undefined;
    const existingPosterPartUri = await this.findByHash(
      resolved.poster.bytes,
      resolved.poster.contentType,
    );
    const slidePart = this.pkg.requirePart(slidePartUri);
    const xml = LosslessXmlDocument.parse(slidePart.bytes);
    const shapeTree = requireMediaShapeTree(xml, slidePartUri);
    const defaultName = `Media ${countDirectMediaPictures(shapeTree)}`;
    const definition = finalizeMediaCreationDefinition(request, resolved, defaultName);

    return this.pkg.transaction(() => {
      const mediaPartUri = definition.media.type === 'embedded'
        ? existingMediaPartUri ?? this.pkg.allocatePartUri(
          '/ppt/media',
          'media',
          definition.media.extension,
        )
        : undefined;
      const posterPartUri = existingPosterPartUri ?? this.pkg.allocatePartUri(
        '/ppt/media',
        'poster',
        definition.poster.extension,
      );
      if (mediaPartUri && definition.media.type === 'embedded' && !this.pkg.hasPart(mediaPartUri)) {
        this.pkg.setPart(mediaPartUri, definition.media.bytes, definition.media.contentType);
      }
      if (!this.pkg.hasPart(posterPartUri)) {
        this.pkg.setPart(posterPartUri, definition.poster.bytes, definition.poster.contentType);
      }

      const kindRelationship = this.pkg.addRelationship(slidePartUri, {
        type: `${REL}${definition.kind}`,
        target: definition.media.type === 'external'
          ? definition.media.url
          : relativeRelationshipTarget(slidePartUri, mediaPartUri!),
        targetMode: definition.media.type === 'external' ? 'External' : 'Internal',
      });
      const mediaRelationship = mediaPartUri
        ? this.pkg.addRelationship(slidePartUri, {
            type: MEDIA_REL,
            target: relativeRelationshipTarget(slidePartUri, mediaPartUri),
            targetMode: 'Internal',
          })
        : undefined;
      const posterRelationship = this.pkg.addRelationship(slidePartUri, {
        type: `${REL}image`,
        target: relativeRelationshipTarget(slidePartUri, posterPartUri),
        targetMode: 'Internal',
      });

      const shapeId = allocateMediaShapeId(xml, slidePartUri);
      const pictureXml = renderMediaPictureXml(shapeId, definition, {
        kind: kindRelationship.id,
        ...(mediaRelationship ? { media: mediaRelationship.id } : {}),
        poster: posterRelationship.id,
      });
      const extensionList = directElementChildren(shapeTree, 'extLst')[0];
      if (extensionList) xml.replace(extensionList.start, extensionList.start, pictureXml);
      else xml.appendChildXml(shapeTree, pictureXml);
      this.pkg.setPart(slidePartUri, xml.serialize(), slidePart.contentType);
      return {
        kind: definition.kind,
        shapeId,
        slidePartUri,
        ...(mediaPartUri ? { mediaPartUri } : {}),
        ...(definition.media.type === 'external' ? { externalUrl: definition.media.url } : {}),
        posterPartUri,
        settings: {
          play: definition.play,
          loop: definition.loop,
          hideWhenStopped: definition.hideWhenStopped,
          volume: definition.volume,
        },
      };
    });
  }

  private editPicture(
    slidePartUri: string,
    shapeId: number,
    edit: (xml: LosslessXmlDocument, picture: XmlElement) => boolean,
  ): void {
    this.pkg.transaction(() => {
      const part = this.pkg.requirePart(slidePartUri);
      const xml = LosslessXmlDocument.parse(part.bytes);
      const picture = xml.elements('pic').find((candidate) => {
        const properties = xml.descendants(candidate, 'cNvPr')[0];
        return Number(properties ? xml.attribute(properties, 'id')?.value : -1) === shapeId;
      });
      if (!picture || !readMediaState(this.pkg, slidePartUri, xml, picture)) {
        throw new Error(`Media shape ${shapeId} was not found on ${slidePartUri}`);
      }
      if (edit(xml, picture)) {
        this.pkg.setPart(slidePartUri, xml.serialize(), part.contentType);
      }
    });
  }

  private async findByHash(bytes: Uint8Array, contentType: string): Promise<string | undefined> {
    const expected = await hash(bytes);
    for (const part of this.pkg.parts) {
      if (
        part.contentType === contentType &&
        part.uri.startsWith('/ppt/media/') &&
        (await hash(part.bytes)) === expected
      ) {
        return part.uri;
      }
    }
    return undefined;
  }
}

async function hash(bytes: Uint8Array): Promise<string> {
  if (!globalThis.crypto?.subtle) throw new Error('Media hashing requires the Web Crypto API');
  const digest = await globalThis.crypto.subtle.digest('SHA-256', new Uint8Array(bytes).buffer);
  return [...new Uint8Array(digest)].map((byte) => byte.toString(16).padStart(2, '0')).join('');
}

function requireMediaShapeTree(
  xml: LosslessXmlDocument,
  slidePartUri: string,
): XmlElement {
  const candidates = xml.elements('spTree').filter((shapeTree) =>
    shapeTree.parent?.localName === 'cSld'
    && shapeTree.parent.parent?.localName === 'sld');
  if (candidates.length !== 1) {
    throw new Error(`Slide ${slidePartUri} must contain exactly one direct shape tree`);
  }
  if (directElementChildren(candidates[0]!, 'extLst').length > 1) {
    throw new Error(`Slide ${slidePartUri} contains repeated shape-tree extension lists`);
  }
  return candidates[0]!;
}

function countDirectMediaPictures(shapeTree: XmlElement): number {
  let count = 0;
  for (const picture of directElementChildren(shapeTree, 'pic')) {
    const nonVisual = directElementChildren(picture, 'nvPicPr')[0];
    const applicationProperties = nonVisual
      ? directElementChildren(nonVisual, 'nvPr')[0]
      : undefined;
    if (
      applicationProperties
      && directElementChildren(applicationProperties, 'audioFile').length
        + directElementChildren(applicationProperties, 'videoFile').length > 0
    ) {
      count += 1;
    }
  }
  return count;
}

function allocateMediaShapeId(
  xml: LosslessXmlDocument,
  slidePartUri: string,
): number {
  let maximum = 1;
  for (const properties of xml.elements('cNvPr')) {
    const value = xml.attribute(properties, 'id')?.value;
    if (value === undefined) continue;
    if (!/^\d+$/.test(value)) {
      throw new Error(`Slide ${slidePartUri} contains an invalid shape id`);
    }
    const id = Number(value);
    if (!Number.isSafeInteger(id) || id > 4_294_967_295) {
      throw new Error(`Slide ${slidePartUri} contains an invalid shape id`);
    }
    maximum = Math.max(maximum, id);
  }
  if (maximum >= 4_294_967_295) {
    throw new Error(`Slide ${slidePartUri} has exhausted its shape ids`);
  }
  return maximum + 1;
}

function directElementChildren(
  element: XmlElement,
  localName: string,
): XmlElement[] {
  return element.children.filter(
    (child): child is XmlElement => child.type === 'element' && child.localName === localName,
  );
}

function ancestor(element: XmlElement, localName: string): XmlElement | undefined {
  let current = element.parent;
  while (current) {
    if (current.localName === localName) return current;
    current = current.parent;
  }
  return undefined;
}
