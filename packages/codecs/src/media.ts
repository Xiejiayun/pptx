import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import { relativeRelationshipTarget, type OpcPackage } from '@pptx/opc';
import type { CodecDiagnostic } from './registry.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const MEDIA_REL = 'http://schemas.microsoft.com/office/2007/relationships/media';
const ONE_PIXEL_PNG = Uint8Array.from([
  137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 1, 0, 0, 0, 1, 8, 4, 0, 0, 0,
  181, 28, 12, 2, 0, 0, 0, 11, 73, 68, 65, 84, 120, 218, 99, 252, 255, 31, 0, 2, 235, 1, 245, 143, 89, 213,
  153, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
]);

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

export interface MediaModel {
  readonly kind: MediaKind;
  readonly shapeId: number;
  readonly slidePartUri: string;
  readonly mediaPartUri?: string;
  readonly externalUrl?: string;
  readonly posterPartUri?: string;
  readonly settings: MediaPlaybackSettings;
}

export class MediaCodec {
  readonly id = 'builtin.media';
  readonly priority = 100;
  readonly ownership = {
    elements: ['a:audioFile', 'a:videoFile', 'p14:media'],
    relationshipTypes: [`${REL}audio`, `${REL}video`, MEDIA_REL],
  } as const;

  constructor(readonly pkg: OpcPackage) {}

  async addAudio(slidePartUri: string, source: MediaSource, options: AddMediaOptions = {}): Promise<MediaModel> {
    return this.add('audio', slidePartUri, source, options);
  }

  async addVideo(slidePartUri: string, source: MediaSource, options: AddMediaOptions = {}): Promise<MediaModel> {
    return this.add('video', slidePartUri, source, options);
  }

  list(slidePartUri: string): readonly MediaModel[] {
    const xml = LosslessXmlDocument.parse(this.pkg.requirePart(slidePartUri).bytes);
    const relationships = this.pkg.relationships(slidePartUri);
    const models: MediaModel[] = [];
    for (const mediaElement of [...xml.elements('audioFile'), ...xml.elements('videoFile')]) {
      const kind: MediaKind = mediaElement.localName === 'audioFile' ? 'audio' : 'video';
      const linkId = xml.attribute(mediaElement, 'r:link')?.value;
      const link = relationships.find(({ id }) => id === linkId);
      const picture = ancestor(mediaElement, 'pic');
      const properties = picture ? xml.descendants(picture, 'cNvPr')[0] : undefined;
      const shapeId = Number(properties ? xml.attribute(properties, 'id')?.value ?? 0 : 0);
      const extensionMedia = picture ? xml.descendants(picture, 'media')[0] : undefined;
      const embedId = extensionMedia ? xml.attribute(extensionMedia, 'r:embed')?.value : undefined;
      const embedded = relationships.find(({ id }) => id === embedId) ?? (link?.targetMode === 'Internal' ? link : undefined);
      const blip = picture ? xml.descendants(picture, 'blip')[0] : undefined;
      const posterId = blip ? xml.attribute(blip, 'r:embed')?.value : undefined;
      const poster = relationships.find(({ id }) => id === posterId);
      const playback = picture ? xml.descendants(picture, 'playback')[0] : undefined;
      models.push({
        kind,
        shapeId,
        slidePartUri,
        ...(embedded?.resolvedTarget ? { mediaPartUri: embedded.resolvedTarget } : {}),
        ...(link?.targetMode === 'External' ? { externalUrl: link.target } : {}),
        ...(poster?.resolvedTarget ? { posterPartUri: poster.resolvedTarget } : {}),
        settings: playback
          ? {
              play: xml.attribute(playback, 'play')?.value === 'auto' ? 'auto' : 'click',
              loop: xml.attribute(playback, 'loop')?.value === '1',
              hideWhenStopped: xml.attribute(playback, 'hideWhenStopped')?.value === '1',
              volume: Number(xml.attribute(playback, 'volume')?.value ?? 100_000) / 100_000,
            }
          : {},
      });
    }
    return models;
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

  diagnostics(model: MediaModel, profile: string): CodecDiagnostic[] {
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
  ): Promise<MediaModel> {
    const resolved = await resolveSource(source, kind, options.contentType, options.fileName);
    let mediaPartUri: string | undefined;
    let externalUrl: string | undefined;
    let embeddedMedia: { readonly bytes: Uint8Array; readonly contentType: string } | undefined;
    let linkTarget: string;
    let linkMode: 'Internal' | 'External';
    if (resolved.externalUrl) {
      externalUrl = resolved.externalUrl;
      linkTarget = resolved.externalUrl;
      linkMode = 'External';
    } else {
      let bytes = resolved.bytes!;
      let contentType = resolved.contentType;
      let extension = resolved.extension;
      if (options.transcode) {
        const transcoded = await options.transcode(bytes, contentType, kind);
        bytes = transcoded.bytes;
        contentType = transcoded.contentType;
        extension = transcoded.extension ?? extensionFor(contentType, kind);
      }
      mediaPartUri =
        (await this.findByHash(bytes, contentType)) ?? this.pkg.allocatePartUri('/ppt/media', 'media', extension);
      embeddedMedia = { bytes, contentType };
      linkTarget = relativeTarget(slidePartUri, mediaPartUri);
      linkMode = 'Internal';
    }

    const poster = await resolvePoster(options.poster, options.posterContentType);
    const posterPartUri =
      (await this.findByHash(poster.bytes, poster.contentType)) ??
      this.pkg.allocatePartUri('/ppt/media', 'poster', poster.extension);
    return this.pkg.transaction(() => {
      if (mediaPartUri && embeddedMedia && !this.pkg.hasPart(mediaPartUri)) {
        this.pkg.setPart(mediaPartUri, embeddedMedia.bytes, embeddedMedia.contentType);
      }
      const link = this.pkg.addRelationship(slidePartUri, {
        type: `${REL}${kind}`,
        target: linkTarget,
        targetMode: linkMode,
      });
      const embedded = mediaPartUri
        ? this.pkg.addRelationship(slidePartUri, {
            type: MEDIA_REL,
            target: relativeTarget(slidePartUri, mediaPartUri),
          })
        : undefined;
      if (!this.pkg.hasPart(posterPartUri)) {
        this.pkg.setPart(posterPartUri, poster.bytes, poster.contentType);
      }
      const posterRelationship = this.pkg.addRelationship(slidePartUri, {
        type: `${REL}image`,
        target: relativeTarget(slidePartUri, posterPartUri),
      });

      const part = this.pkg.requirePart(slidePartUri);
      const xml = LosslessXmlDocument.parse(part.bytes);
      const shapeTree = xml.elements('spTree')[0];
      if (!shapeTree) throw new Error(`Slide ${slidePartUri} has no shape tree`);
      const shapeId =
        Math.max(1, ...xml.elements('cNvPr').map((element) => Number(xml.attribute(element, 'id')?.value ?? 0))) + 1;
      const position = {
        x: options.x ?? 914_400,
        y: options.y ?? 914_400,
        width: options.width ?? (kind === 'video' ? 4_572_000 : 914_400),
        height: options.height ?? (kind === 'video' ? 2_571_750 : 914_400),
      };
      const mediaExtension = embedded
        ? `<p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"><p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="${embedded.id}"/></p:ext>`
        : '';
      const playbackExtension = `<p:ext uri="{C13D3E4A-5148-4B6D-A7E7-505054582D4F}"><px:playback xmlns:px="urn:pptx-ooxml:media" play="${
        options.play ?? 'click'
      }" loop="${options.loop ? 1 : 0}" hideWhenStopped="${options.hideWhenStopped ? 1 : 0}" volume="${Math.round(
        clamp(options.volume ?? 1) * 100_000,
      )}"/></p:ext>`;
      const extension = `<p:extLst>${mediaExtension}${playbackExtension}</p:extLst>`;
      const picture = `<p:pic><p:nvPicPr><p:cNvPr id="${shapeId}" name="${kind === 'video' ? 'Video' : 'Audio'} ${shapeId}"><a:hlinkClick r:id="" action="ppaction://media"/></p:cNvPr><p:cNvPicPr/><p:nvPr><a:${kind}File r:link="${link.id}"/>${extension}</p:nvPr></p:nvPicPr><p:blipFill><a:blip r:embed="${posterRelationship.id}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${position.x}" y="${position.y}"/><a:ext cx="${position.width}" cy="${position.height}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
      xml.appendChildXml(shapeTree, picture);
      this.pkg.setPart(slidePartUri, xml.serialize(), part.contentType);
      return {
        kind,
        shapeId,
        slidePartUri,
        ...(mediaPartUri ? { mediaPartUri } : {}),
        ...(externalUrl ? { externalUrl } : {}),
        posterPartUri,
        settings: {
          play: options.play ?? 'click',
          loop: options.loop ?? false,
          hideWhenStopped: options.hideWhenStopped ?? false,
          volume: clamp(options.volume ?? 1),
        },
      };
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

async function resolveSource(
  source: MediaSource,
  kind: MediaKind,
  contentType?: string,
  fileName?: string,
): Promise<{ bytes?: Uint8Array; externalUrl?: string; contentType: string; extension: string }> {
  if (typeof source === 'string' && /^https?:\/\//i.test(source)) {
    return { externalUrl: source, contentType: contentType ?? defaultContentType(kind), extension: extensionFor(contentType ?? defaultContentType(kind), kind) };
  }
  let bytes: Uint8Array;
  let name = fileName;
  if (typeof source === 'string') {
    const fs = await loadNodeModule<NodeFsPromises>(['node:fs', 'promises'].join('/'));
    bytes = new Uint8Array(await fs.readFile(source));
    name = name ?? source;
  } else if (source instanceof Uint8Array) bytes = new Uint8Array(source);
  else if (source instanceof ArrayBuffer) bytes = new Uint8Array(source);
  else if (isBlob(source)) {
    bytes = new Uint8Array(await source.arrayBuffer());
    const sourceName = (source as Blob & { readonly name?: unknown }).name;
    if (!name && typeof sourceName === 'string') name = sourceName;
  } else bytes = await readStream(source);
  const sourceExtension = name ? fileExtension(name) : '';
  const inferred = contentType ?? contentTypeFor(sourceExtension, kind);
  return {
    bytes,
    contentType: inferred,
    extension: sourceExtension || extensionFor(inferred, kind),
  };
}

async function resolvePoster(source?: MediaSource, contentType?: string): Promise<{ bytes: Uint8Array; contentType: string; extension: string }> {
  if (!source || (typeof source === 'string' && /^https?:\/\//i.test(source))) {
    return { bytes: ONE_PIXEL_PNG, contentType: 'image/png', extension: '.png' };
  }
  const resolved = await resolveSource(source, 'video', contentType ?? 'image/png');
  return { bytes: resolved.bytes!, contentType: resolved.contentType, extension: resolved.extension };
}

async function readStream(stream: MediaByteStream): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  if (isReadableStream(stream)) {
    const reader = stream.getReader();
    try {
      while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        chunks.push(normalizeByteChunk(value));
      }
    } finally {
      reader.releaseLock();
    }
  } else {
    for await (const chunk of stream) chunks.push(normalizeByteChunk(chunk));
  }
  return concatenateBytes(chunks);
}

function defaultContentType(kind: MediaKind): string {
  return kind === 'video' ? 'video/mp4' : 'audio/mpeg';
}

function contentTypeFor(extension: string, kind: MediaKind): string {
  return {
    '.mp4': 'video/mp4',
    '.mov': 'video/quicktime',
    '.webm': 'video/webm',
    '.mp3': 'audio/mpeg',
    '.m4a': 'audio/mp4',
    '.wav': 'audio/wav',
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.jpeg': 'image/jpeg',
  }[extension.toLowerCase()] ?? defaultContentType(kind);
}

function extensionFor(contentType: string, kind: MediaKind): string {
  return {
    'video/mp4': '.mp4',
    'video/quicktime': '.mov',
    'video/webm': '.webm',
    'audio/mpeg': '.mp3',
    'audio/mp4': '.m4a',
    'audio/wav': '.wav',
    'image/png': '.png',
    'image/jpeg': '.jpg',
  }[contentType] ?? (kind === 'video' ? '.mp4' : '.mp3');
}

function relativeTarget(sourcePartUri: string, targetPartUri: string): string {
  return relativeRelationshipTarget(sourcePartUri, targetPartUri);
}

function fileExtension(value: string): string {
  const basename = value.replaceAll('\\', '/').split('/').at(-1) ?? '';
  const dot = basename.lastIndexOf('.');
  return dot <= 0 ? '' : basename.slice(dot);
}

async function hash(bytes: Uint8Array): Promise<string> {
  if (!globalThis.crypto?.subtle) throw new Error('Media hashing requires the Web Crypto API');
  const digest = await globalThis.crypto.subtle.digest('SHA-256', new Uint8Array(bytes).buffer);
  return [...new Uint8Array(digest)].map((byte) => byte.toString(16).padStart(2, '0')).join('');
}

interface NodeFsPromises {
  readFile(path: string): Promise<Uint8Array>;
}

async function loadNodeModule<T>(specifier: string): Promise<T> {
  if (typeof process === 'undefined' || !process.versions?.node) {
    throw new Error('Local media paths are only supported in Node.js; pass a Blob, File, or byte stream');
  }
  return import(specifier) as Promise<T>;
}

function isReadableStream(value: unknown): value is ReadableStream<unknown> {
  return Boolean(value && typeof (value as { getReader?: unknown }).getReader === 'function');
}

function isBlob(value: unknown): value is Blob {
  return typeof Blob !== 'undefined' && value instanceof Blob;
}

function normalizeByteChunk(chunk: unknown): Uint8Array {
  if (typeof chunk === 'number' && Number.isInteger(chunk) && chunk >= 0 && chunk <= 255) {
    return Uint8Array.of(chunk);
  }
  if (chunk instanceof Uint8Array) return new Uint8Array(chunk);
  if (chunk instanceof ArrayBuffer) return new Uint8Array(chunk);
  if (ArrayBuffer.isView(chunk)) {
    return new Uint8Array(new Uint8Array(chunk.buffer, chunk.byteOffset, chunk.byteLength));
  }
  throw new TypeError('Media streams must yield byte numbers, Uint8Array, ArrayBuffer, or ArrayBufferView chunks');
}

function concatenateBytes(chunks: readonly Uint8Array[]): Uint8Array {
  const output = new Uint8Array(chunks.reduce((total, chunk) => total + chunk.byteLength, 0));
  let offset = 0;
  for (const chunk of chunks) {
    output.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return output;
}

function ancestor(element: XmlElement, localName: string): XmlElement | undefined {
  let current = element.parent;
  while (current) {
    if (current.localName === localName) return current;
    current = current.parent;
  }
  return undefined;
}

function clamp(value: number): number {
  return Math.max(0, Math.min(1, Number.isFinite(value) ? value : 1));
}
