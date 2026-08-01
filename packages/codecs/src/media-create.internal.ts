import { escapeXmlAttribute } from '@pptx/lossless-xml';
import type {
  AddMediaOptions,
  MediaByteStream,
  MediaKind,
  MediaPlaceholderIdentity,
  MediaPlaceholderSelector,
} from './media.js';
import type {
  ResolvedEmbeddedMedia,
  ResolvedExternalMedia,
  ResolvedMediaCreationInputs,
} from './media-source.internal.js';

const OFFICE_MEDIA_EXTENSION_URI = '{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}';
const PLAYBACK_EXTENSION_URI = '{C13D3E4A-5148-4B6D-A7E7-505054582D4F}';

const OPTION_KEYS = new Set([
  'name',
  'altText',
  'placeholder',
  'contentType',
  'fileName',
  'poster',
  'posterContentType',
  'x',
  'y',
  'width',
  'height',
  'play',
  'loop',
  'hideWhenStopped',
  'volume',
  'transcode',
]);

export type NormalizedMediaSourceReference =
  | Readonly<{ type: 'string'; value: string }>
  | Readonly<{ type: 'bytes'; bytes: Uint8Array }>
  | Readonly<{ type: 'blob'; value: Blob }>
  | Readonly<{ type: 'stream'; value: MediaByteStream }>;

export interface NormalizedMediaCreateRequest {
  readonly kind: MediaKind;
  readonly source: NormalizedMediaSourceReference;
  readonly poster?: NormalizedMediaSourceReference;
  readonly name?: string;
  readonly altText?: string;
  readonly placeholder?: MediaPlaceholderSelector;
  readonly contentType?: string;
  readonly posterContentType?: string;
  readonly fileName?: string;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly play: 'click' | 'auto';
  readonly loop: boolean;
  readonly hideWhenStopped: boolean;
  readonly volume: number;
  readonly transcode?: NonNullable<AddMediaOptions['transcode']>;
}

export interface NormalizedMediaCreationDefinition {
  readonly kind: MediaKind;
  readonly name: string;
  readonly altText?: string;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly rotation: number;
  readonly flipHorizontal: boolean;
  readonly flipVertical: boolean;
  readonly play: 'click' | 'auto';
  readonly loop: boolean;
  readonly hideWhenStopped: boolean;
  readonly volume: number;
  readonly media: ResolvedEmbeddedMedia | ResolvedExternalMedia;
  readonly poster: ResolvedEmbeddedMedia;
}

export interface MediaRelationshipIds {
  readonly kind: string;
  readonly media?: string;
  readonly poster: string;
}

export function normalizeMediaCreateRequest(
  kind: unknown,
  source: unknown,
  options: unknown,
): Readonly<NormalizedMediaCreateRequest> {
  const normalizedKind = normalizeKind(kind);
  const values = readOptions(options);
  const poster = values.poster === undefined
    ? undefined
    : normalizeSource(values.poster, 'poster');
  const name = normalizeXmlString(values.name, 'name');
  const altText = normalizeXmlString(values.altText, 'altText');
  const placeholder = values.placeholder === undefined
    ? undefined
    : normalizePlaceholderSelector(values.placeholder);
  const contentType = normalizeNonEmptyString(values.contentType, 'contentType');
  const posterContentType = normalizeNonEmptyString(
    values.posterContentType,
    'posterContentType',
  );
  const fileName = normalizeNonEmptyString(values.fileName, 'fileName');
  const transcode = normalizeTranscode(values.transcode);

  return Object.freeze({
    kind: normalizedKind,
    source: normalizeSource(source, 'source'),
    ...(poster === undefined ? {} : { poster }),
    ...(name === undefined ? {} : { name }),
    ...(altText === undefined ? {} : { altText }),
    ...(placeholder === undefined ? {} : { placeholder }),
    ...(contentType === undefined ? {} : { contentType }),
    ...(posterContentType === undefined ? {} : { posterContentType }),
    ...(fileName === undefined ? {} : { fileName }),
    x: normalizeInteger(values.x, 914_400, 'x'),
    y: normalizeInteger(values.y, 914_400, 'y'),
    width: normalizePositiveInteger(
      values.width,
      normalizedKind === 'video' ? 4_572_000 : 914_400,
      'width',
    ),
    height: normalizePositiveInteger(
      values.height,
      normalizedKind === 'video' ? 2_571_750 : 914_400,
      'height',
    ),
    play: normalizePlay(values.play),
    loop: normalizeBoolean(values.loop, false, 'loop'),
    hideWhenStopped: normalizeBoolean(
      values.hideWhenStopped,
      false,
      'hideWhenStopped',
    ),
    volume: normalizeVolume(values.volume),
    ...(transcode === undefined ? {} : { transcode }),
  });
}

export function finalizeMediaCreationDefinition(
  request: Readonly<NormalizedMediaCreateRequest>,
  resolved: Readonly<ResolvedMediaCreationInputs>,
  defaultName: string,
): Readonly<NormalizedMediaCreationDefinition> {
  const name = request.name ?? normalizeRequiredXmlString(defaultName, 'default name');
  const media = resolved.media.type === 'embedded'
    ? cloneEmbeddedMedia(resolved.media)
    : Object.freeze({ type: 'external' as const, url: resolved.media.url });
  return Object.freeze({
    kind: request.kind,
    name,
    ...(request.altText === undefined ? {} : { altText: request.altText }),
    x: request.x,
    y: request.y,
    width: request.width,
    height: request.height,
    rotation: 0,
    flipHorizontal: false,
    flipVertical: false,
    play: request.play,
    loop: request.loop,
    hideWhenStopped: request.hideWhenStopped,
    volume: request.volume,
    media,
    poster: cloneEmbeddedMedia(resolved.poster),
  });
}

export function renderMediaPictureXml(
  shapeId: number,
  definition: Readonly<NormalizedMediaCreationDefinition>,
  relationships: Readonly<MediaRelationshipIds>,
  placeholder?: Readonly<MediaPlaceholderIdentity>,
): string {
  if (!Number.isSafeInteger(shapeId) || shapeId <= 0) {
    throw new RangeError('Media shape id must be a positive safe integer');
  }
  const kindRelationshipId = normalizeRelationshipId(relationships.kind, 'kind');
  const posterRelationshipId = normalizeRelationshipId(relationships.poster, 'poster');
  let mediaExtension = '';
  if (definition.media.type === 'embedded') {
    const mediaRelationshipId = normalizeRelationshipId(relationships.media, 'media');
    mediaExtension = `<p:ext uri="${OFFICE_MEDIA_EXTENSION_URI}">`
      + '<p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" '
      + `r:embed="${escapeXmlAttribute(mediaRelationshipId)}"/></p:ext>`;
  } else if (relationships.media !== undefined) {
    throw new TypeError('External media must not have an embedded media relationship');
  }

  const name = escapeXmlAttribute(definition.name);
  const description = definition.altText === undefined
    ? ''
    : ` descr="${escapeXmlAttribute(definition.altText)}"`;
  const playbackExtension = `<p:ext uri="${PLAYBACK_EXTENSION_URI}">`
    + '<px:playback xmlns:px="urn:pptx-ooxml:media" '
    + `play="${definition.play}" loop="${definition.loop ? 1 : 0}" `
    + `hideWhenStopped="${definition.hideWhenStopped ? 1 : 0}" `
    + `volume="${Math.round(definition.volume * 100_000)}"/></p:ext>`;
  const placeholderXml = placeholder === undefined
    ? ''
    : `<p:ph type="${placeholder.type}" idx="${placeholder.index}"/>`;
  const transformAttributes = [
    definition.rotation === 0 ? '' : ` rot="${definition.rotation}"`,
    definition.flipHorizontal ? ' flipH="1"' : '',
    definition.flipVertical ? ' flipV="1"' : '',
  ].join('');

  return '<p:pic><p:nvPicPr>'
    + `<p:cNvPr id="${shapeId}" name="${name}"${description}>`
    + '<a:hlinkClick r:id="" action="ppaction://media"/></p:cNvPr>'
    + `<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr>${placeholderXml}`
    + `<a:${definition.kind}File r:link="${escapeXmlAttribute(kindRelationshipId)}"/>`
    + `<p:extLst>${mediaExtension}${playbackExtension}</p:extLst></p:nvPr></p:nvPicPr>`
    + `<p:blipFill><a:blip r:embed="${escapeXmlAttribute(posterRelationshipId)}"/>`
    + `<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm${transformAttributes}>`
    + `<a:off x="${definition.x}" y="${definition.y}"/>`
    + `<a:ext cx="${definition.width}" cy="${definition.height}"/></a:xfrm>`
    + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>';
}

function normalizePlaceholderSelector(value: unknown): MediaPlaceholderSelector {
  if (typeof value === 'string') {
    if (value.length === 0) throw new TypeError('Media placeholder name must not be empty');
    if (!isValidXmlString(value)) {
      throw new TypeError('Media placeholder name contains invalid XML characters');
    }
    return value;
  }
  const identity = readPlaceholderIdentity(value);
  const types = ['title', 'body', 'pic', 'chart', 'tbl', 'media'] as const;
  if (!types.includes(identity.type as typeof types[number])) {
    throw new TypeError(`Media placeholder type must be ${types.join(', ')}`);
  }
  if (
    typeof identity.index !== 'number'
    || !Number.isSafeInteger(identity.index)
    || identity.index < 0
    || identity.index > 4_294_967_294
  ) {
    throw new RangeError('Media placeholder index must be between 0 and 4294967294');
  }
  return Object.freeze({
    type: identity.type as MediaPlaceholderIdentity['type'],
    index: identity.index,
  });
}

function readPlaceholderIdentity(value: unknown): Record<'type' | 'index', unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError('Media placeholder must be an object');
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError('Media placeholder must be an ordinary object');
  }
  const result = Object.create(null) as Record<string, unknown>;
  for (const key of Reflect.ownKeys(value)) {
    if (key !== 'type' && key !== 'index') {
      throw new TypeError(`Media placeholder contains unsupported property ${String(key)}`);
    }
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`Media placeholder ${String(key)} must be a data property`);
    }
    result[key] = descriptor.value;
  }
  if (!Object.hasOwn(result, 'type') || !Object.hasOwn(result, 'index')) {
    throw new TypeError('Media placeholder requires type and index');
  }
  return result as Record<'type' | 'index', unknown>;
}

function cloneEmbeddedMedia(
  value: Readonly<ResolvedEmbeddedMedia>,
): ResolvedEmbeddedMedia {
  return Object.freeze({
    type: 'embedded',
    bytes: new Uint8Array(value.bytes),
    contentType: value.contentType,
    extension: value.extension,
  });
}

function normalizeRequiredXmlString(value: unknown, name: string): string {
  if (typeof value !== 'string') throw new TypeError(`Media ${name} must be a string`);
  if (!isValidXmlString(value)) {
    throw new TypeError(`Media ${name} contains invalid XML characters`);
  }
  return value;
}

function normalizeRelationshipId(value: unknown, name: string): string {
  const normalized = normalizeRequiredXmlString(value, `${name} relationship id`);
  if (normalized.length === 0) {
    throw new TypeError(`Media ${name} relationship id must be a non-empty string`);
  }
  return normalized;
}

function readOptions(value: unknown): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError('Media creation options must be an object');
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError('Media creation options must be an ordinary object');
  }
  const result = Object.create(null) as Record<string, unknown>;
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !OPTION_KEYS.has(key)) {
      throw new TypeError(`Media creation options contain unsupported property ${String(key)}`);
    }
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`Media creation option ${key} must be a data property`);
    }
    result[key] = descriptor.value;
  }
  return result;
}

function normalizeKind(value: unknown): MediaKind {
  if (value === 'audio' || value === 'video') return value;
  throw new TypeError('Media kind must be audio or video');
}

function normalizeSource(value: unknown, label: string): NormalizedMediaSourceReference {
  if (typeof value === 'string') {
    if (value.length === 0) throw new TypeError(`Media ${label} string must not be empty`);
    return Object.freeze({ type: 'string', value });
  }
  if (value instanceof Uint8Array) {
    if (value.length === 0) throw new RangeError(`Media ${label} bytes must not be empty`);
    return Object.freeze({ type: 'bytes', bytes: new Uint8Array(value) });
  }
  if (value instanceof ArrayBuffer) {
    if (value.byteLength === 0) throw new RangeError(`Media ${label} bytes must not be empty`);
    return Object.freeze({
      type: 'bytes',
      bytes: new Uint8Array(new Uint8Array(value)),
    });
  }
  if (isBlob(value)) return Object.freeze({ type: 'blob', value });
  if (isMediaByteStream(value)) {
    return Object.freeze({ type: 'stream', value });
  }
  throw new TypeError(`Media ${label} must be a path, URL, byte buffer, Blob, or byte stream`);
}

function normalizeXmlString(value: unknown, name: string): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'string') throw new TypeError(`Media ${name} must be a string`);
  if (!isValidXmlString(value)) {
    throw new TypeError(`Media ${name} contains invalid XML characters`);
  }
  return value;
}

function normalizeNonEmptyString(value: unknown, name: string): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'string' || value.length === 0) {
    throw new TypeError(`Media ${name} must be a non-empty string`);
  }
  return value;
}

function normalizeInteger(value: unknown, defaultValue: number, name: string): number {
  if (value === undefined) return defaultValue;
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`Media ${name} must be finite`);
  }
  if (!Number.isSafeInteger(value)) {
    throw new RangeError(`Media ${name} must be a safe integer`);
  }
  return value === 0 ? 0 : value;
}

function normalizePositiveInteger(value: unknown, defaultValue: number, name: string): number {
  const normalized = normalizeInteger(value, defaultValue, name);
  if (normalized <= 0) throw new RangeError(`Media ${name} must be positive`);
  return normalized;
}

function normalizePlay(value: unknown): 'click' | 'auto' {
  if (value === undefined) return 'click';
  if (value === 'click' || value === 'auto') return value;
  throw new TypeError('Media play must be click or auto');
}

function normalizeBoolean(value: unknown, defaultValue: boolean, name: string): boolean {
  if (value === undefined) return defaultValue;
  if (typeof value !== 'boolean') throw new TypeError(`Media ${name} must be a boolean`);
  return value;
}

function normalizeVolume(value: unknown): number {
  if (value === undefined) return 1;
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError('Media volume must be finite');
  }
  if (value < 0 || value > 1) throw new RangeError('Media volume must be between 0 and 1');
  return value === 0 ? 0 : value;
}

function normalizeTranscode(
  value: unknown,
): NonNullable<AddMediaOptions['transcode']> | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'function') throw new TypeError('Media transcode must be a function');
  return value as NonNullable<AddMediaOptions['transcode']>;
}

function isBlob(value: unknown): value is Blob {
  return typeof Blob !== 'undefined' && value instanceof Blob;
}

function isMediaByteStream(value: unknown): value is MediaByteStream {
  return Boolean(
    value
    && (
      typeof (value as { getReader?: unknown }).getReader === 'function'
      || typeof (value as { [Symbol.asyncIterator]?: unknown })[Symbol.asyncIterator] === 'function'
    ),
  );
}

function isValidXmlString(value: string): boolean {
  for (const character of value) {
    const codePoint = character.codePointAt(0)!;
    if (
      codePoint === 0x09
      || codePoint === 0x0A
      || codePoint === 0x0D
      || (codePoint >= 0x20 && codePoint <= 0xD7FF)
      || (codePoint >= 0xE000 && codePoint <= 0xFFFD)
      || (codePoint >= 0x10000 && codePoint <= 0x10FFFF)
    ) continue;
    return false;
  }
  return true;
}
