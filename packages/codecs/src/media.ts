import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import { relativeRelationshipTarget, type OpcPackage } from '@pptx/opc';
import {
  finalizeMediaCreationDefinition,
  normalizeMediaCreateRequest,
  renderMediaPictureXml,
} from './media-create.internal.js';
import {
  normalizeMediaAltText,
  normalizeMediaName,
  normalizeMediaPlaybackSettings,
  readMediaPlaybackExtension,
  replaceMediaMetadataAttribute,
  replaceMediaPlaybackExtension,
} from './media-edit.internal.js';
import { resolveMediaCreationInputs } from './media-source.internal.js';
import { readMediaState } from './media-state.internal.js';
import {
  findMatchingMediaPart,
  normalizeMediaPosterReplaceRequest,
  normalizeMediaReplaceRequest,
  replaceResolvedMediaPoster,
  replaceResolvedMediaSource,
  resolveMediaReplacementPoster,
  resolveMediaReplacementSource,
} from './media-replace.internal.js';
import {
  clearNativeMediaTiming,
  syncNativeMediaTiming,
} from './media-timing-edit.internal.js';
import { readNativeMediaTiming } from './media-timing-state.internal.js';
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

export interface ReplaceMediaSourceOptions {
  readonly contentType?: string;
  readonly fileName?: string;
  readonly transcode?: AddMediaOptions['transcode'];
}

export interface ReplaceMediaPosterOptions {
  readonly contentType?: string;
  readonly fileName?: string;
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
      const preference = readMediaPlaybackExtension(xml, picture);
      if (preference.malformed) {
        throw new Error(`Media shape ${shapeId} has a malformed playback extension`);
      }
      if (normalized) {
        const sync = syncNativeMediaTiming(
          xml,
          shapeId,
          state.kind,
          normalized,
          preference.ownership,
        );
        const preferenceChanged = replaceMediaPlaybackExtension(
          xml,
          picture,
          normalized,
          sync.ownership,
        );
        return sync.changed || preferenceChanged;
      }
      const timingChanged = clearNativeMediaTiming(
        xml,
        shapeId,
        state.kind,
        preference.ownership,
      );
      return replaceMediaPlaybackExtension(xml, picture, undefined) || timingChanged;
    });
  }

  async replaceSource(
    slidePartUri: string,
    shapeId: number,
    kind: MediaKind,
    source: MediaSource,
    options: ReplaceMediaSourceOptions = {},
  ): Promise<MediaDescriptor> {
    const request = normalizeMediaReplaceRequest(kind, source, options);
    const resolved = await resolveMediaReplacementSource(request);
    const matching = resolved.type === 'embedded'
      ? await findMatchingMediaPart(this.pkg, resolved)
      : undefined;
    replaceResolvedMediaSource(
      this.pkg,
      slidePartUri,
      shapeId,
      kind,
      resolved,
      matching,
    );
    const descriptor = this.list(slidePartUri).find((candidate) => candidate.shapeId === shapeId);
    if (!descriptor) throw new Error(`Media shape ${shapeId} was not found on ${slidePartUri}`);
    return descriptor;
  }

  async replacePoster(
    slidePartUri: string,
    shapeId: number,
    source?: MediaSource,
    options: ReplaceMediaPosterOptions = {},
  ): Promise<MediaDescriptor> {
    const request = normalizeMediaPosterReplaceRequest(source, options);
    const resolved = await resolveMediaReplacementPoster(request);
    const matching = await findMatchingMediaPart(this.pkg, resolved);
    replaceResolvedMediaPoster(this.pkg, slidePartUri, shapeId, resolved, matching);
    const descriptor = this.list(slidePartUri).find((candidate) => candidate.shapeId === shapeId);
    if (!descriptor) throw new Error(`Media shape ${shapeId} was not found on ${slidePartUri}`);
    return descriptor;
  }

  delete(slidePartUri: string, shapeId: number): void {
    this.pkg.transaction(() => {
      const part = this.pkg.requirePart(slidePartUri);
      const xml = LosslessXmlDocument.parse(part.bytes);
      const picture = xml.elements('pic').find((candidate) => {
        const properties = xml.descendants(candidate, 'cNvPr')[0];
        return Number(properties ? xml.attribute(properties, 'id')?.value : -1) === shapeId;
      });
      const state = picture ? readMediaState(this.pkg, slidePartUri, xml, picture) : undefined;
      if (!picture || !state) {
        throw new Error(`Media shape ${shapeId} was not found on ${slidePartUri}`);
      }
      const preference = readMediaPlaybackExtension(xml, picture);
      const ids = new Set(
        xml
          .descendants(picture)
          .flatMap((element) => element.attributes)
          .filter(({ name, value }) => name.startsWith('r:') && value.length > 0)
          .map(({ value }) => value),
      );
      const targetIds = mediaRelationshipIds(xml, picture);
      const targets = this.pkg
        .relationships(slidePartUri)
        .filter(({ id, type }) => targetIds.has(id) && (
          type === `${REL}audio`
          || type === `${REL}video`
          || type === `${REL}image`
          || type === MEDIA_REL
        ))
        .map(({ resolvedTarget }) => resolvedTarget)
        .filter((target): target is string => Boolean(target));
      clearNativeMediaTiming(xml, shapeId, state.kind, preference.ownership);
      xml.removeElement(picture);
      const updated = xml.serialize();
      this.pkg.setPart(slidePartUri, updated, part.contentType);
      const remaining = LosslessXmlDocument.parse(updated);
      for (const id of ids) {
        if (relationshipReferenceCount(remaining, id) === 0) {
          this.pkg.removeRelationship(slidePartUri, id);
        }
      }
      for (const target of new Set(targets)) {
        const incoming = this.pkg.graph.find(({ uri }) => uri === target)?.incoming ?? [];
        if (incoming.length === 0 && target.startsWith('/ppt/media/')) this.pkg.deletePart(target);
      }
    });
  }

  materializePlayback(slidePartUri: string, shapeId?: number): number {
    return this.pkg.transaction(() => {
      const part = this.pkg.requirePart(slidePartUri);
      let xml = LosslessXmlDocument.parse(part.bytes);
      const shapeIds = directMediaPictures(xml, shapeId)
        .map((picture) => mediaPictureShapeId(xml, picture))
        .filter((id): id is number => id !== undefined);
      let changed = 0;
      for (const id of shapeIds) {
        const picture = directMediaPictures(xml, id)[0]!;
        const state = readMediaState(this.pkg, slidePartUri, xml, picture);
        const preference = readMediaPlaybackExtension(xml, picture);
        if (!state || !preference.settings || preference.malformed) continue;
        const sync = syncNativeMediaTiming(
          xml,
          state.shapeId,
          state.kind,
          preference.settings,
          preference.ownership,
        );
        const preferenceChanged = replaceMediaPlaybackExtension(
          xml,
          picture,
          preference.settings,
          sync.ownership,
        );
        if (sync.changed || preferenceChanged) {
          changed += 1;
          xml = LosslessXmlDocument.parse(xml.serialize());
        }
      }
      if (changed > 0) this.pkg.setPart(slidePartUri, xml.serialize(), part.contentType);
      return changed;
    });
  }

  diagnostics(model: MediaDescriptor, profile: string): CodecDiagnostic[] {
    const xml = LosslessXmlDocument.parse(this.pkg.requirePart(model.slidePartUri).bytes);
    const slideModels = xml.elements('pic')
      .map((picture) => readMediaState(this.pkg, model.slidePartUri, xml, picture))
      .filter((candidate): candidate is NonNullable<typeof candidate> => candidate !== undefined);
    return [
      ...mediaProfileDiagnostics(this.pkg, model, profile),
      ...nativeMediaTimingDiagnostics(this.pkg, model.slidePartUri, xml, slideModels, [model]),
    ];
  }

  diagnosticsForSlide(slidePartUri: string, profile: string): CodecDiagnostic[] {
    const xml = LosslessXmlDocument.parse(this.pkg.requirePart(slidePartUri).bytes);
    const models = xml.elements('pic')
      .map((picture) => readMediaState(this.pkg, slidePartUri, xml, picture))
      .filter((model): model is NonNullable<typeof model> => model !== undefined);
    return [
      ...models.flatMap((model) => mediaProfileDiagnostics(this.pkg, model, profile)),
      ...nativeMediaTimingDiagnostics(this.pkg, slidePartUri, xml, models),
    ];
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
      const updatedXml = LosslessXmlDocument.parse(xml.serialize());
      const picture = updatedXml.elements('pic').find((candidate) => {
        const properties = updatedXml.descendants(candidate, 'cNvPr')[0];
        return Number(properties ? updatedXml.attribute(properties, 'id')?.value : -1) === shapeId;
      });
      if (!picture) throw new Error(`Created media shape ${shapeId} is missing`);
      const sync = syncNativeMediaTiming(
        updatedXml,
        shapeId,
        definition.kind,
        definition,
      );
      replaceMediaPlaybackExtension(updatedXml, picture, definition, sync.ownership);
      this.pkg.setPart(slidePartUri, updatedXml.serialize(), slidePart.contentType);
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

function mediaProfileDiagnostics(
  pkg: OpcPackage,
  model: MediaDescriptor,
  profile: string,
): CodecDiagnostic[] {
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
  if (model.mediaPartUri) {
    const contentType = pkg.getPart(model.mediaPartUri)?.contentType;
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

interface MediaTimingTargetIssue {
  readonly diagnostic: CodecDiagnostic;
  readonly shapeId: number;
  readonly mediaTnId?: number;
}

function nativeMediaTimingDiagnostics(
  pkg: OpcPackage,
  slidePartUri: string,
  xml: LosslessXmlDocument,
  slideModels: readonly MediaDescriptor[],
  diagnosedModels: readonly MediaDescriptor[] = slideModels,
): CodecDiagnostic[] {
  const targetIssues = mediaTimingTargetIssues(xml, slidePartUri, slideModels);
  const diagnostics = targetIssues.map(({ diagnostic }) => diagnostic);
  for (const model of diagnosedModels) {
    const picture = xml.elements('pic').find((candidate) => {
      const properties = xml.descendants(candidate, 'cNvPr')[0];
      return Number(properties ? xml.attribute(properties, 'id')?.value : -1) === model.shapeId;
    });
    if (!picture || !readMediaState(pkg, slidePartUri, xml, picture)) continue;
    const preference = readMediaPlaybackExtension(xml, picture);
    const native = readNativeMediaTiming(
      xml,
      model.shapeId,
      model.kind,
      preference.ownership,
    );
    const hasTargetIssue = targetIssues.some((issue) =>
      issue.shapeId === model.shapeId
      || (
        issue.mediaTnId !== undefined
        && issue.mediaTnId === preference.ownership?.mediaTnId
      ));
    if (hasTargetIssue) continue;

    if (preference.malformed) {
      diagnostics.push(timingDiagnostic(
        'MEDIA_TIMING_STALE',
        `Media shape ${model.shapeId} has malformed playback ownership`,
        slidePartUri,
      ));
      continue;
    }
    if (native.status === 'unsupported') {
      diagnostics.push(timingDiagnostic(
        'MEDIA_TIMING_UNSUPPORTED',
        `Media shape ${model.shapeId} uses unsupported native timing: ${native.reason ?? 'unknown structure'}`,
        slidePartUri,
      ));
      continue;
    }
    if (native.status === 'ambiguous') {
      diagnostics.push(timingDiagnostic(
        'MEDIA_TIMING_AMBIGUOUS',
        `Media shape ${model.shapeId} has ambiguous native timing: ${native.reason ?? 'unknown structure'}`,
        slidePartUri,
      ));
      continue;
    }
    if (
      native.status === 'absent'
      || (native.status === 'owned-stale' && native.reason?.includes('missing'))
    ) {
      if (preference.settings) {
        diagnostics.push(timingDiagnostic(
          'MEDIA_TIMING_MISSING',
          `Media shape ${model.shapeId} has playback preferences but no native timing graph`,
          slidePartUri,
        ));
      }
      continue;
    }
    if (
      native.status === 'owned-stale'
      || (native.status === 'recognized-imported' && Boolean(preference.settings))
      || (preference.settings && native.settings
        && !mediaTimingSettingsEqual(preference.settings, native.settings))
    ) {
      diagnostics.push(timingDiagnostic(
        'MEDIA_TIMING_STALE',
        `Media shape ${model.shapeId} playback ownership or settings do not match native timing`,
        slidePartUri,
      ));
    }
  }
  return diagnostics;
}

function mediaTimingTargetIssues(
  xml: LosslessXmlDocument,
  slidePartUri: string,
  models: readonly MediaDescriptor[],
): readonly MediaTimingTargetIssue[] {
  const modelsByShapeId = new Map<number, MediaDescriptor>();
  for (const model of models) modelsByShapeId.set(model.shapeId, model);
  const issues: MediaTimingTargetIssue[] = [];
  const seen = new Set<string>();
  for (const mediaNode of xml.elements('cMediaNode')) {
    if (!ancestor(mediaNode, 'timing')) continue;
    const mediaKind = mediaNode.parent?.localName === 'audio' || mediaNode.parent?.localName === 'video'
      ? mediaNode.parent.localName
      : undefined;
    const mediaTnId = directElementChildren(mediaNode, 'cTn').length === 1
      ? positiveInteger(xml.attribute(directElementChildren(mediaNode, 'cTn')[0]!, 'id')?.value)
      : undefined;
    for (const target of xml.descendants(mediaNode, 'spTgt')) {
      const shapeId = positiveInteger(xml.attribute(target, 'spid')?.value);
      if (shapeId === undefined) continue;
      const model = modelsByShapeId.get(shapeId);
      const code = !model
        ? 'MEDIA_TIMING_DANGLING_TARGET'
        : mediaKind && mediaKind !== model.kind
          ? 'MEDIA_TIMING_KIND_MISMATCH'
          : undefined;
      if (!code || seen.has(`${code}:${shapeId}`)) continue;
      seen.add(`${code}:${shapeId}`);
      issues.push({
        diagnostic: timingDiagnostic(
          code,
          code === 'MEDIA_TIMING_DANGLING_TARGET'
            ? `Native media timing targets missing media shape ${shapeId}`
            : `Native ${mediaKind} timing does not match media shape ${shapeId} kind ${model!.kind}`,
          slidePartUri,
        ),
        shapeId,
        ...(mediaTnId === undefined ? {} : { mediaTnId }),
      });
    }
  }
  return issues;
}

function mediaTimingSettingsEqual(
  left: Readonly<Required<MediaPlaybackSettings>>,
  right: Readonly<Required<MediaPlaybackSettings>>,
): boolean {
  return left.play === right.play
    && left.loop === right.loop
    && left.hideWhenStopped === right.hideWhenStopped
    && left.volume === right.volume;
}

function timingDiagnostic(
  code: string,
  message: string,
  partUri: string,
): CodecDiagnostic {
  return { severity: 'warning', code, message, partUri };
}

function positiveInteger(value: string | undefined): number | undefined {
  if (!value || !/^[1-9]\d*$/.test(value)) return undefined;
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) && parsed <= 0xFFFF_FFFF ? parsed : undefined;
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

function directMediaPictures(
  xml: LosslessXmlDocument,
  shapeId?: number,
): readonly XmlElement[] {
  const slides = xml.roots.filter(({ localName }) => localName === 'sld');
  const commonSlides = slides.length === 1
    ? directElementChildren(slides[0]!, 'cSld')
    : [];
  const shapeTrees = commonSlides.length === 1
    ? directElementChildren(commonSlides[0]!, 'spTree')
    : [];
  if (shapeTrees.length !== 1) return [];
  return directElementChildren(shapeTrees[0]!, 'pic').filter((picture) => {
    const value = mediaPictureShapeId(xml, picture);
    return value !== undefined && (shapeId === undefined || value === shapeId);
  });
}

function mediaPictureShapeId(
  xml: LosslessXmlDocument,
  picture: XmlElement,
): number | undefined {
  const nonVisual = directElementChildren(picture, 'nvPicPr');
  const properties = nonVisual.length === 1
    ? directElementChildren(nonVisual[0]!, 'cNvPr')
    : [];
  return properties.length === 1
    ? positiveInteger(xml.attribute(properties[0]!, 'id')?.value)
    : undefined;
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

function relationshipReferenceCount(xml: LosslessXmlDocument, id: string): number {
  return xml.elements().flatMap(({ attributes }) => attributes)
    .filter(({ name, value }) => name.startsWith('r:') && value === id).length;
}

function mediaRelationshipIds(xml: LosslessXmlDocument, picture: XmlElement): ReadonlySet<string> {
  const ids = new Set<string>();
  for (const element of xml.descendants(picture)) {
    let attribute: 'r:embed' | 'r:link' | undefined;
    if (element.localName === 'audioFile' || element.localName === 'videoFile') {
      attribute = 'r:link';
    } else if (
      element.localName === 'media'
      || (element.localName === 'blip' && element.parent?.localName === 'blipFill')
    ) {
      attribute = 'r:embed';
    }
    const id = attribute ? xml.attribute(element, attribute)?.value : undefined;
    if (id) ids.add(id);
  }
  return ids;
}
