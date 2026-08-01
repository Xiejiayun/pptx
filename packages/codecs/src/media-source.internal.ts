import type { MediaByteChunk, MediaByteStream, MediaKind } from './media.js';
import type {
  NormalizedMediaCreateRequest,
  NormalizedMediaSourceReference,
} from './media-create.internal.js';

const ONE_PIXEL_PNG = Uint8Array.from([
  137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 1,
  0, 0, 0, 1, 8, 4, 0, 0, 0, 181, 28, 12, 2, 0, 0, 0, 11, 73, 68, 65, 84,
  120, 218, 99, 252, 255, 31, 0, 2, 235, 1, 245, 143, 89, 213, 153, 0, 0, 0,
  0, 73, 69, 78, 68, 174, 66, 96, 130,
]);

const CONTENT_TYPE_EXTENSIONS: Readonly<Record<string, readonly string[]>> = {
  'audio/mpeg': ['.mp3'],
  'audio/mp4': ['.m4a'],
  'audio/wav': ['.wav'],
  'audio/ogg': ['.ogg'],
  'video/mp4': ['.mp4', '.m4v'],
  'video/quicktime': ['.mov'],
  'video/webm': ['.webm'],
  'image/png': ['.png'],
  'image/jpeg': ['.jpg', '.jpeg'],
  'image/gif': ['.gif'],
};

const EXTENSION_CONTENT_TYPES = Object.freeze(
  Object.fromEntries(
    Object.entries(CONTENT_TYPE_EXTENSIONS).flatMap(([contentType, extensions]) =>
      extensions.map((extension) => [extension, contentType])),
  ) as Readonly<Record<string, string>>,
);

const DOMAIN_CONTENT_TYPES: Readonly<Record<MediaDomain, ReadonlySet<string>>> = {
  audio: new Set(['audio/mpeg', 'audio/mp4', 'audio/wav', 'audio/ogg']),
  video: new Set(['video/mp4', 'video/quicktime', 'video/webm']),
  poster: new Set(['image/png', 'image/jpeg', 'image/gif']),
};

const TRANSCODE_RESULT_KEYS = new Set(['bytes', 'contentType', 'extension']);

type MediaDomain = MediaKind | 'poster';

interface PreparedMediaSource {
  readonly reference: NormalizedMediaSourceReference;
  readonly dataUri?: Readonly<{ bytes: Uint8Array; contentType: string }>;
  readonly externalUrl?: string;
  readonly name?: string;
}

interface ResolvedMediaDescriptor {
  readonly contentType: string;
  readonly extension: string;
}

export interface ResolvedEmbeddedMedia {
  readonly type: 'embedded';
  readonly bytes: Uint8Array;
  readonly contentType: string;
  readonly extension: string;
}

export interface ResolvedExternalMedia {
  readonly type: 'external';
  readonly url: string;
}

export interface ResolvedMediaCreationInputs {
  readonly media: ResolvedEmbeddedMedia | ResolvedExternalMedia;
  readonly poster: ResolvedEmbeddedMedia;
}

export async function resolveMediaCreationInputs(
  request: Readonly<NormalizedMediaCreateRequest>,
): Promise<Readonly<ResolvedMediaCreationInputs>> {
  const preparedMedia = prepareSource(request.source, 'media', true);
  const preparedPoster = request.poster
    ? prepareSource(request.poster, 'poster', false)
    : undefined;
  if (preparedMedia.externalUrl && request.transcode) {
    throw new TypeError('External media cannot be transcoded');
  }

  const mediaDescriptor = resolveDescriptor(
    request.kind,
    request.contentType,
    preparedMedia.dataUri?.contentType,
    request.fileName ?? preparedMedia.name,
  );
  const posterDescriptor = resolveDescriptor(
    'poster',
    request.posterContentType,
    preparedPoster?.dataUri?.contentType ?? (preparedPoster ? undefined : 'image/png'),
    preparedPoster?.name,
  );

  let media: ResolvedEmbeddedMedia | ResolvedExternalMedia;
  if (preparedMedia.externalUrl) {
    media = Object.freeze({ type: 'external', url: preparedMedia.externalUrl });
  } else {
    media = createEmbedded(await loadPreparedBytes(preparedMedia, 'media'), mediaDescriptor);
    if (request.transcode) media = await applyTranscode(media, request);
  }
  const poster = preparedPoster
    ? createEmbedded(await loadPreparedBytes(preparedPoster, 'poster'), posterDescriptor)
    : createEmbedded(ONE_PIXEL_PNG, posterDescriptor);
  return Object.freeze({ media, poster });
}

function prepareSource(
  reference: NormalizedMediaSourceReference,
  label: 'media' | 'poster',
  allowExternal: boolean,
): Readonly<PreparedMediaSource> {
  if (reference.type === 'string') {
    if (reference.value.startsWith('data:')) {
      return Object.freeze({ reference, dataUri: parseDataUri(reference.value) });
    }
    if (/^https?:\/\//i.test(reference.value)) {
      validateExternalUrl(reference.value);
      if (!allowExternal) throw new TypeError('External poster URLs are not supported');
      const url = new URL(reference.value);
      return Object.freeze({
        reference,
        externalUrl: reference.value,
        name: url.pathname,
      });
    }
    if (hasUriScheme(reference.value) && !isWindowsDrivePath(reference.value)) {
      throw new TypeError(`Unsupported ${label} URL scheme`);
    }
    return Object.freeze({ reference, name: reference.value });
  }
  if (reference.type === 'blob') {
    if (nativeBlobSize(reference.value) === 0) {
      throw new RangeError(`Media ${label} Blob must not be empty`);
    }
    const name = nativeFileName(reference.value);
    return Object.freeze({ reference, ...(name === undefined ? {} : { name }) });
  }
  return Object.freeze({ reference });
}

function resolveDescriptor(
  domain: MediaDomain,
  assertedContentType: string | undefined,
  declaredContentType: string | undefined,
  name: string | undefined,
): Readonly<ResolvedMediaDescriptor> {
  if (assertedContentType !== undefined) validateContentType(assertedContentType, domain);
  if (declaredContentType !== undefined) validateContentType(declaredContentType, domain);
  if (
    assertedContentType !== undefined
    && declaredContentType !== undefined
    && assertedContentType !== declaredContentType
  ) {
    throw new TypeError(`${domain} content type assertion does not match the data URI`);
  }

  const hintedExtension = name === undefined ? undefined : fileExtension(name);
  const hintedContentType = hintedExtension
    ? EXTENSION_CONTENT_TYPES[hintedExtension]
    : undefined;
  if (hintedContentType !== undefined) validateContentType(hintedContentType, domain);
  const resolvedContentType = assertedContentType
    ?? declaredContentType
    ?? hintedContentType
    ?? defaultContentType(domain);
  if (hintedContentType !== undefined && hintedContentType !== resolvedContentType) {
    throw new TypeError(`${domain} extension does not match content type ${resolvedContentType}`);
  }
  return Object.freeze({
    contentType: resolvedContentType,
    extension: hintedContentType === undefined
      ? canonicalExtension(resolvedContentType)
      : hintedExtension!,
  });
}

async function loadPreparedBytes(
  prepared: Readonly<PreparedMediaSource>,
  label: 'media' | 'poster',
): Promise<Uint8Array> {
  if (prepared.dataUri) return new Uint8Array(prepared.dataUri.bytes);
  const { reference } = prepared;
  let bytes: Uint8Array;
  if (reference.type === 'bytes') {
    bytes = new Uint8Array(reference.bytes);
  } else if (reference.type === 'blob') {
    bytes = new Uint8Array(await reference.value.arrayBuffer());
  } else if (reference.type === 'stream') {
    bytes = await readStream(reference.value);
  } else {
    const fs = await loadNodeModule<NodeFsPromises>(['node:fs', 'promises'].join('/'));
    bytes = new Uint8Array(await fs.readFile(reference.value));
  }
  if (bytes.length === 0) throw new RangeError(`Resolved ${label} payload is empty`);
  return bytes;
}

async function applyTranscode(
  media: Readonly<ResolvedEmbeddedMedia>,
  request: Readonly<NormalizedMediaCreateRequest>,
): Promise<ResolvedEmbeddedMedia> {
  const result = await request.transcode!(
    new Uint8Array(media.bytes),
    media.contentType,
    request.kind,
  );
  const values = readTranscodeResult(result);
  if (!(values.bytes instanceof Uint8Array)) {
    throw new TypeError('Media transcode bytes must be a Uint8Array');
  }
  if (values.bytes.length === 0) throw new RangeError('Media transcode bytes must not be empty');
  if (typeof values.contentType !== 'string') {
    throw new TypeError('Media transcode contentType must be a string');
  }
  validateContentType(values.contentType, request.kind);
  const extension = values.extension === undefined
    ? canonicalExtension(values.contentType)
    : normalizeTranscodeExtension(values.extension, values.contentType);
  return createEmbedded(values.bytes, { contentType: values.contentType, extension });
}

function readTranscodeResult(value: unknown): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError('Media transcode result must be an object');
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError('Media transcode result must be an ordinary object');
  }
  const result = Object.create(null) as Record<string, unknown>;
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !TRANSCODE_RESULT_KEYS.has(key)) {
      throw new TypeError(`Media transcode result contains unsupported property ${String(key)}`);
    }
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`Media transcode result ${key} must be a data property`);
    }
    result[key] = descriptor.value;
  }
  return result;
}

function normalizeTranscodeExtension(value: unknown, contentType: string): string {
  if (typeof value !== 'string' || value.length === 0) {
    throw new TypeError('Media transcode extension must be a non-empty string');
  }
  if (value !== value.toLowerCase()) {
    throw new TypeError('Media transcode extension must be lowercase');
  }
  const extension = value.startsWith('.') ? value : `.${value}`;
  if (!/^\.[a-z0-9]+$/.test(extension)) {
    throw new TypeError('Media transcode extension is invalid');
  }
  const extensionContentType = EXTENSION_CONTENT_TYPES[extension];
  if (extensionContentType === undefined || extensionContentType !== contentType) {
    throw new TypeError('Media transcode extension does not match its content type');
  }
  return extension;
}

function createEmbedded(
  bytes: Uint8Array,
  descriptor: Readonly<ResolvedMediaDescriptor>,
): ResolvedEmbeddedMedia {
  if (bytes.length === 0) throw new RangeError('Embedded media bytes must not be empty');
  return Object.freeze({
    type: 'embedded',
    bytes: new Uint8Array(bytes),
    contentType: descriptor.contentType,
    extension: descriptor.extension,
  });
}

function parseDataUri(value: string): Readonly<{ bytes: Uint8Array; contentType: string }> {
  const match = /^data:([^;,]+);base64,([A-Za-z0-9+/]*={0,2})$/.exec(value);
  if (!match || match[0] !== value) {
    throw new TypeError('Media data URI must contain a supported MIME type and canonical base64 data');
  }
  const contentType = match[1]!;
  if (!Object.hasOwn(CONTENT_TYPE_EXTENSIONS, contentType)) {
    throw new TypeError(`Unsupported media data URI content type ${contentType}`);
  }
  return Object.freeze({
    bytes: decodeCanonicalBase64(match[2]!),
    contentType,
  });
}

function decodeCanonicalBase64(payload: string): Uint8Array {
  if (payload.length === 0) throw new RangeError('Media data URI payload is empty');
  if (payload.length % 4 !== 0) {
    throw new TypeError('Media data URI payload has invalid base64 length');
  }
  const firstPadding = payload.indexOf('=');
  const padding = firstPadding === -1 ? 0 : payload.length - firstPadding;
  if (padding > 2) throw new TypeError('Media data URI payload has invalid base64 padding');
  const dataLength = payload.length - padding;
  for (let index = 0; index < dataLength; index += 1) {
    if (base64Value(payload.charCodeAt(index)) < 0) {
      throw new TypeError('Media data URI payload contains an invalid base64 character');
    }
  }
  for (let index = dataLength; index < payload.length; index += 1) {
    if (payload[index] !== '=') {
      throw new TypeError('Media data URI payload has invalid base64 padding');
    }
  }
  if (padding === 2 && (base64Value(payload.charCodeAt(payload.length - 3)) & 0x0f) !== 0) {
    throw new TypeError('Media data URI payload has non-canonical base64 padding bits');
  }
  if (padding === 1 && (base64Value(payload.charCodeAt(payload.length - 2)) & 0x03) !== 0) {
    throw new TypeError('Media data URI payload has non-canonical base64 padding bits');
  }

  const output = new Uint8Array((payload.length / 4) * 3 - padding);
  let outputOffset = 0;
  for (let inputOffset = 0; inputOffset < payload.length; inputOffset += 4) {
    const first = base64Value(payload.charCodeAt(inputOffset));
    const second = base64Value(payload.charCodeAt(inputOffset + 1));
    const third = payload[inputOffset + 2] === '='
      ? 0
      : base64Value(payload.charCodeAt(inputOffset + 2));
    const fourth = payload[inputOffset + 3] === '='
      ? 0
      : base64Value(payload.charCodeAt(inputOffset + 3));
    output[outputOffset] = (first << 2) | (second >>> 4);
    outputOffset += 1;
    if (outputOffset < output.length) {
      output[outputOffset] = ((second & 0x0f) << 4) | (third >>> 2);
      outputOffset += 1;
    }
    if (outputOffset < output.length) {
      output[outputOffset] = ((third & 0x03) << 6) | fourth;
      outputOffset += 1;
    }
  }
  return output;
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

function normalizeByteChunk(chunk: unknown): Uint8Array {
  if (typeof chunk === 'number' && Number.isInteger(chunk) && chunk >= 0 && chunk <= 255) {
    return Uint8Array.of(chunk);
  }
  if (chunk instanceof Uint8Array) return new Uint8Array(chunk);
  if (chunk instanceof ArrayBuffer) return new Uint8Array(new Uint8Array(chunk));
  if (ArrayBuffer.isView(chunk)) {
    return new Uint8Array(new Uint8Array(chunk.buffer, chunk.byteOffset, chunk.byteLength));
  }
  throw new TypeError('Media streams must yield byte numbers or ArrayBuffer byte views');
}

function concatenateBytes(chunks: readonly Uint8Array[]): Uint8Array {
  const output = new Uint8Array(chunks.reduce((total, chunk) => total + chunk.length, 0));
  let offset = 0;
  for (const chunk of chunks) {
    output.set(chunk, offset);
    offset += chunk.length;
  }
  return output;
}

function validateContentType(contentType: string, domain: MediaDomain): void {
  if (!DOMAIN_CONTENT_TYPES[domain].has(contentType)) {
    throw new TypeError(`Unsupported ${domain} content type ${contentType}`);
  }
}

function defaultContentType(domain: MediaDomain): string {
  if (domain === 'audio') return 'audio/mpeg';
  if (domain === 'video') return 'video/mp4';
  return 'image/png';
}

function canonicalExtension(contentType: string): string {
  return CONTENT_TYPE_EXTENSIONS[contentType]![0]!;
}

function fileExtension(value: string): string | undefined {
  const basename = value.replaceAll('\\', '/').split('/').at(-1) ?? '';
  const dot = basename.lastIndexOf('.');
  return dot <= 0 ? undefined : basename.slice(dot).toLowerCase();
}

function validateExternalUrl(value: string): void {
  let parsed: URL;
  try {
    parsed = new URL(value);
  } catch {
    throw new TypeError('External media URL is invalid');
  }
  if ((parsed.protocol !== 'http:' && parsed.protocol !== 'https:') || parsed.hostname.length === 0) {
    throw new TypeError('External media URL must use HTTP or HTTPS');
  }
}

function nativeFileName(value: Blob): string | undefined {
  if (typeof File === 'undefined' || !(value instanceof File)) return undefined;
  const getter = Object.getOwnPropertyDescriptor(File.prototype, 'name')?.get;
  if (!getter) return undefined;
  const name = getter.call(value) as unknown;
  return typeof name === 'string' && name.length > 0 ? name : undefined;
}

function nativeBlobSize(value: Blob): number {
  const getter = Object.getOwnPropertyDescriptor(Blob.prototype, 'size')?.get;
  if (!getter) return value.size;
  return getter.call(value) as number;
}

function hasUriScheme(value: string): boolean {
  return /^[A-Za-z][A-Za-z\d+.-]*:/.test(value);
}

function isWindowsDrivePath(value: string): boolean {
  return /^[A-Za-z]:[\\/]/.test(value);
}

function isReadableStream(value: MediaByteStream): value is ReadableStream<MediaByteChunk> {
  return typeof (value as { getReader?: unknown }).getReader === 'function';
}

function base64Value(code: number): number {
  if (code >= 65 && code <= 90) return code - 65;
  if (code >= 97 && code <= 122) return code - 71;
  if (code >= 48 && code <= 57) return code + 4;
  if (code === 43) return 62;
  if (code === 47) return 63;
  return -1;
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
