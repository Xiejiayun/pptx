import type {
  AddSvgImageOptions,
  RasterImageContentType,
  SvgImageContentType,
} from '@pptx/model';
import {
  normalizeImageSizing,
  type ImageSizing,
} from './raster-image-sizing.js';
import { inspectSvgImage, type SvgImageInfo } from './svg-image-source.js';

const PNG_SIGNATURE = Uint8Array.from([137, 80, 78, 71, 13, 10, 26, 10]);
const GIF87A_SIGNATURE = Uint8Array.from([71, 73, 70, 56, 55, 97]);
const GIF89A_SIGNATURE = Uint8Array.from([71, 73, 70, 56, 57, 97]);
const JPEG_SOF_MARKERS = new Set([
  0xc0, 0xc1, 0xc2, 0xc3,
  0xc5, 0xc6, 0xc7,
  0xc9, 0xca, 0xcb,
  0xcd, 0xce, 0xcf,
]);
const MODEL_IMAGE_OPTION_KEYS = new Set([
  'name',
  'altText',
  'placeholder',
  'x',
  'y',
  'width',
  'height',
  'rotation',
  'flipHorizontal',
  'flipVertical',
]);
const ADD_IMAGE_SOURCE_OPTION_KEYS = new Set([
  'contentType',
  'fallback',
  'signal',
  'sizing',
  ...MODEL_IMAGE_OPTION_KEYS,
]);

export interface RasterImageInfo {
  readonly contentType: RasterImageContentType;
  readonly width: number;
  readonly height: number;
}

type AddImageSourceBaseOptions = Omit<
  AddSvgImageOptions,
  'sourceRectangle' | 'width' | 'height'
> & {
  readonly contentType?: ImageContentType;
  readonly fallback?: ImageSource;
  readonly signal?: AbortSignal;
};

export type AddImageSourceOptions = AddImageSourceBaseOptions & (
  | {
      readonly sizing?: undefined;
      readonly width?: number;
      readonly height?: number;
    }
  | {
      readonly sizing: ImageSizing;
      readonly width?: never;
      readonly height?: never;
    }
);

export interface NormalizedAddImageSourceOptions {
  readonly contentType?: ImageContentType;
  readonly fallback?: ImageSource;
  readonly signal?: AbortSignal;
  readonly imageOptions: Readonly<
    Omit<AddSvgImageOptions, 'sourceRectangle'>
  >;
  readonly sizing?: Readonly<ImageSizing>;
}

export type RasterImageByteChunk = number | Uint8Array | ArrayBuffer | ArrayBufferView;
export type RasterImageByteStream =
  | ReadableStream<RasterImageByteChunk>
  | AsyncIterable<RasterImageByteChunk>;
export type RasterImageSource =
  | string
  | Uint8Array
  | ArrayBuffer
  | Blob
  | RasterImageByteStream;

export interface ResolvedRasterImageSource {
  readonly bytes: Uint8Array;
  readonly info: RasterImageInfo;
  readonly assertedContentType?: RasterImageContentType;
}

export type ImageContentType = RasterImageContentType | SvgImageContentType;
export type ImageByteChunk = RasterImageByteChunk;
export type ImageByteStream = RasterImageByteStream;
export type ImageSource = RasterImageSource;
export type ImageInfo = RasterImageInfo | SvgImageInfo;

export interface ResolvedImageSource {
  readonly bytes: Uint8Array;
  readonly info: ImageInfo;
  readonly assertedContentType?: ImageContentType;
}

interface LoadedImageSource {
  readonly bytes: Uint8Array;
  readonly assertedContentType?: ImageContentType;
}

export function inspectRasterImage(bytes: Uint8Array): RasterImageInfo {
  if (!(bytes instanceof Uint8Array)) {
    throw new TypeError('Raster image bytes must be a Uint8Array');
  }
  if (bytes.byteLength === 0) throw new TypeError('Raster image bytes cannot be empty');
  if (startsWith(bytes, PNG_SIGNATURE)) return inspectPng(bytes);
  if (startsWith(bytes, GIF87A_SIGNATURE) || startsWith(bytes, GIF89A_SIGNATURE)) {
    return inspectGif(bytes);
  }
  if (bytes[0] === 0xff && bytes[1] === 0xd8) return inspectJpeg(bytes);
  if (isTruncatedPrefix(bytes, PNG_SIGNATURE)) {
    throw new TypeError('Truncated PNG signature');
  }
  if (
    isTruncatedPrefix(bytes, GIF87A_SIGNATURE) ||
    isTruncatedPrefix(bytes, GIF89A_SIGNATURE)
  ) {
    throw new TypeError('Truncated GIF signature');
  }
  if (bytes.byteLength === 1 && bytes[0] === 0xff) {
    throw new TypeError('Truncated JPEG signature');
  }
  throw new TypeError('Unsupported raster image signature');
}

export function inspectImage(bytes: Uint8Array): ImageInfo {
  if (!(bytes instanceof Uint8Array)) {
    throw new TypeError('Image bytes must be a Uint8Array');
  }
  if (bytes.byteLength === 0) throw new TypeError('Image bytes cannot be empty');
  return isRasterImageCandidate(bytes)
    ? inspectRasterImage(bytes)
    : inspectSvgImage(bytes);
}

export function normalizeAddImageSourceOptions(
  options: unknown,
): NormalizedAddImageSourceOptions {
  if (!options || typeof options !== 'object' || Array.isArray(options)) {
    throw new TypeError('Raster image source options must be an object');
  }
  const prototype = Object.getPrototypeOf(options);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError('Raster image source options must be an ordinary object');
  }
  const values = Object.create(null) as Record<string, unknown>;
  for (const key of Reflect.ownKeys(options)) {
    if (typeof key !== 'string' || !ADD_IMAGE_SOURCE_OPTION_KEYS.has(key)) {
      throw new TypeError(
        `Raster image source options contain unsupported property ${String(key)}`,
      );
    }
    const descriptor = Object.getOwnPropertyDescriptor(options, key);
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`Raster image source option ${key} must be a data property`);
    }
    values[key] = descriptor.value;
  }

  const contentType = normalizeOptionalContentType(values.contentType);
  const fallback = values.fallback === undefined
    ? undefined
    : normalizeImageSourceOption(values.fallback, 'SVG fallback');
  const signal = normalizeOptionalAbortSignal(values.signal);
  const sizing = values.sizing === undefined
    ? undefined
    : normalizeImageSizing(values.sizing);
  if (
    sizing !== undefined
    && (Object.hasOwn(values, 'width') || Object.hasOwn(values, 'height'))
  ) {
    throw new TypeError('Raster image sizing cannot be combined with top-level width or height');
  }
  const imageOptions = Object.create(null) as Record<string, unknown>;
  for (const key of MODEL_IMAGE_OPTION_KEYS) {
    if (Object.hasOwn(values, key)) imageOptions[key] = values[key];
  }
  Object.freeze(imageOptions);
  return Object.freeze({
    imageOptions: imageOptions as Omit<AddSvgImageOptions, 'sourceRectangle'>,
    ...(contentType === undefined ? {} : { contentType }),
    ...(fallback === undefined ? {} : { fallback }),
    ...(signal === undefined ? {} : { signal }),
    ...(sizing === undefined ? {} : { sizing }),
  });
}

export function assertRasterImageContentType(
  expected: RasterImageContentType | undefined,
  resolved: Readonly<ResolvedRasterImageSource>,
): void {
  if (expected !== undefined && expected !== resolved.info.contentType) {
    throw new TypeError(
      `Raster image contentType expected ${expected} but the signature is ${resolved.info.contentType}`,
    );
  }
  if (
    resolved.assertedContentType !== undefined &&
    resolved.assertedContentType !== resolved.info.contentType
  ) {
    throw new TypeError(
      `Raster image source declares ${resolved.assertedContentType} but the signature is ${resolved.info.contentType}`,
    );
  }
}

export function assertImageContentType(
  expected: ImageContentType | undefined,
  resolved: Readonly<ResolvedImageSource>,
): void {
  if (expected !== undefined && expected !== resolved.info.contentType) {
    throw new TypeError(
      `Image contentType expected ${expected} but the detected type is ${resolved.info.contentType}`,
    );
  }
  if (
    resolved.assertedContentType !== undefined
    && resolved.assertedContentType !== resolved.info.contentType
  ) {
    throw new TypeError(
      `Image source declares ${resolved.assertedContentType} but the detected type is ${resolved.info.contentType}`,
    );
  }
}

export async function resolveRasterImageSource(
  source: RasterImageSource,
  signal?: AbortSignal,
): Promise<ResolvedRasterImageSource> {
  const loaded = await loadImageSource(source, signal);
  if (loaded.assertedContentType === 'image/svg+xml') {
    throw new TypeError('Raster image data URI cannot declare image/svg+xml');
  }
  const resolved: ResolvedRasterImageSource = {
    bytes: loaded.bytes,
    info: inspectRasterImage(loaded.bytes),
    ...(loaded.assertedContentType === undefined
      ? {}
      : { assertedContentType: loaded.assertedContentType }),
  };
  assertRasterImageContentType(undefined, resolved);
  return resolved;
}

export async function resolveImageSource(
  source: ImageSource,
  signal?: AbortSignal,
): Promise<ResolvedImageSource> {
  const loaded = await loadImageSource(source, signal);
  const resolved: ResolvedImageSource = {
    bytes: loaded.bytes,
    info: inspectImage(loaded.bytes),
    ...(loaded.assertedContentType === undefined
      ? {}
      : { assertedContentType: loaded.assertedContentType }),
  };
  assertImageContentType(undefined, resolved);
  return resolved;
}

async function loadImageSource(
  source: ImageSource,
  signal?: AbortSignal,
): Promise<LoadedImageSource> {
  throwIfAborted(signal);
  if (typeof source === 'string') {
    const resolved = await resolveStringSource(source, signal);
    throwIfAborted(signal);
    return resolved;
  }
  let bytes: Uint8Array;
  if (source instanceof Uint8Array) {
    bytes = new Uint8Array(source);
  } else if (source instanceof ArrayBuffer) {
    bytes = new Uint8Array(new Uint8Array(source));
  } else if (isBlob(source)) {
    const buffer = await source.arrayBuffer();
    throwIfAborted(signal);
    bytes = new Uint8Array(new Uint8Array(buffer));
  } else if (isReadableStream(source)) {
    bytes = await readWebStream(source, signal);
  } else if (isAsyncIterable(source)) {
    bytes = await readAsyncIterable(source, signal);
  } else {
    throw new TypeError('Unsupported image source type');
  }
  throwIfAborted(signal);
  return { bytes };
}

async function resolveStringSource(
  source: string,
  signal?: AbortSignal,
): Promise<LoadedImageSource> {
  if (source.length === 0) throw new TypeError('Image source string cannot be empty');
  if (/^data:/i.test(source)) return resolveDataUri(source);
  if (/^https?:\/\//i.test(source)) return resolveFetchSource(source, signal);
  if (hasUriScheme(source) && !isWindowsDrivePath(source)) {
    throw new TypeError('Unsupported image URL scheme');
  }
  return isNodeRuntime()
    ? resolveNodePath(source, signal)
    : resolveFetchSource(source, signal);
}

async function resolveNodePath(
  path: string,
  signal?: AbortSignal,
): Promise<LoadedImageSource> {
  throwIfAborted(signal);
  const fs = await loadNodeModule<NodeFsPromises>(['node:fs', 'promises'].join('/'));
  let input: Uint8Array;
  try {
    input = await fs.readFile(path, signal ? { signal } : undefined);
  } catch (error) {
    throwIfAborted(signal);
    throw error;
  }
  throwIfAborted(signal);
  const bytes = new Uint8Array(input);
  return { bytes };
}

async function resolveFetchSource(
  url: string,
  signal?: AbortSignal,
): Promise<LoadedImageSource> {
  if (typeof globalThis.fetch !== 'function') {
    throw new Error('Image URL loading requires the Fetch API');
  }
  throwIfAborted(signal);
  let response: Response;
  try {
    response = await globalThis.fetch(url, signal ? { signal } : undefined);
  } catch (error) {
    throwIfAborted(signal);
    throw error;
  }
  throwIfAborted(signal);
  if (!response.ok) {
    throw new TypeError(`Image request failed with HTTP ${response.status}`);
  }
  let input: ArrayBuffer;
  try {
    input = await response.arrayBuffer();
  } catch (error) {
    throwIfAborted(signal);
    throw error;
  }
  throwIfAborted(signal);
  const bytes = new Uint8Array(new Uint8Array(input));
  return { bytes };
}

function resolveDataUri(source: string): LoadedImageSource {
  const match = /^data:(image\/(?:png|jpeg|gif|svg\+xml));base64,([A-Za-z0-9+/]*={0,2})$/i.exec(source);
  if (!match || match[0] !== source) {
    throw new TypeError(
      'Image data URI must use image/png, image/jpeg, image/gif, or image/svg+xml with canonical base64 data',
    );
  }
  const assertedContentType = match[1]!.toLowerCase() as ImageContentType;
  const bytes = decodeCanonicalBase64(match[2]!);
  return { bytes, assertedContentType };
}

function decodeCanonicalBase64(payload: string): Uint8Array {
  if (payload.length === 0) throw new TypeError('Image data URI payload cannot be empty');
  if (payload.length % 4 !== 0) {
    throw new TypeError('Image data URI payload has invalid base64 length');
  }
  const firstPadding = payload.indexOf('=');
  const padding = firstPadding === -1 ? 0 : payload.length - firstPadding;
  if (padding > 2) throw new TypeError('Image data URI payload has invalid base64 padding');
  const dataLength = payload.length - padding;
  for (let index = 0; index < dataLength; index += 1) {
    if (base64Value(payload.charCodeAt(index)) < 0) {
      throw new TypeError('Image data URI payload contains an invalid base64 character');
    }
  }
  for (let index = dataLength; index < payload.length; index += 1) {
    if (payload[index] !== '=') {
      throw new TypeError('Image data URI payload has invalid base64 padding');
    }
  }
  if (padding === 2 && (base64Value(payload.charCodeAt(payload.length - 3)) & 0x0f) !== 0) {
    throw new TypeError('Image data URI payload has non-canonical base64 padding bits');
  }
  if (padding === 1 && (base64Value(payload.charCodeAt(payload.length - 2)) & 0x03) !== 0) {
    throw new TypeError('Image data URI payload has non-canonical base64 padding bits');
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

function base64Value(code: number): number {
  if (code >= 65 && code <= 90) return code - 65;
  if (code >= 97 && code <= 122) return code - 71;
  if (code >= 48 && code <= 57) return code + 4;
  if (code === 43) return 62;
  if (code === 47) return 63;
  return -1;
}

function normalizeOptionalContentType(value: unknown): ImageContentType | undefined {
  if (value === undefined) return undefined;
  if (
    value === 'image/png'
    || value === 'image/jpeg'
    || value === 'image/gif'
    || value === 'image/svg+xml'
  ) return value;
  throw new TypeError('Image source contentType is unsupported');
}

function normalizeImageSourceOption(value: unknown, context: string): ImageSource {
  if (typeof value === 'string') {
    if (value.length === 0) throw new TypeError(`${context} source string cannot be empty`);
    return value;
  }
  if (value instanceof Uint8Array) return new Uint8Array(value);
  if (value instanceof ArrayBuffer) return value.slice(0);
  if (
    isBlob(value)
    || hasDataMethod(value, 'getReader')
    || hasDataMethod(value, Symbol.asyncIterator)
  ) return value as Blob | RasterImageByteStream;
  throw new TypeError(`${context} source type is unsupported`);
}

function hasDataMethod(value: unknown, key: PropertyKey): boolean {
  if (!value || (typeof value !== 'object' && typeof value !== 'function')) return false;
  let current: object | null = value as object;
  while (current) {
    const descriptor = Object.getOwnPropertyDescriptor(current, key);
    if (descriptor) return Object.hasOwn(descriptor, 'value')
      && typeof descriptor.value === 'function';
    current = Object.getPrototypeOf(current) as object | null;
  }
  return false;
}

function normalizeOptionalAbortSignal(value: unknown): AbortSignal | undefined {
  if (value === undefined) return undefined;
  if (typeof AbortSignal !== 'undefined' && value instanceof AbortSignal) return value;
  throw new TypeError('Raster image source signal must be an AbortSignal');
}

interface NodeFsPromises {
  readFile(path: string, options?: { readonly signal?: AbortSignal }): Promise<Uint8Array>;
}

async function loadNodeModule<T>(specifier: string): Promise<T> {
  return import(specifier) as Promise<T>;
}

function isNodeRuntime(): boolean {
  return typeof process !== 'undefined' && Boolean(process.versions?.node);
}

function hasUriScheme(value: string): boolean {
  return /^[A-Za-z][A-Za-z0-9+.-]*:/.test(value);
}

function isWindowsDrivePath(value: string): boolean {
  return /^[A-Za-z]:/.test(value);
}

function inspectPng(bytes: Uint8Array): RasterImageInfo {
  if (bytes.byteLength < 24) throw new TypeError('Truncated PNG IHDR');
  if (readUint32BigEndian(bytes, 8) !== 13) {
    throw new TypeError('PNG IHDR length must be 13');
  }
  if (bytes[12] !== 73 || bytes[13] !== 72 || bytes[14] !== 68 || bytes[15] !== 82) {
    throw new TypeError('PNG first chunk must be IHDR');
  }
  const width = readUint32BigEndian(bytes, 16);
  const height = readUint32BigEndian(bytes, 20);
  if (width === 0 || height === 0) throw new TypeError('PNG dimensions must be positive');
  return Object.freeze({ contentType: 'image/png', width, height });
}

function inspectGif(bytes: Uint8Array): RasterImageInfo {
  if (bytes.byteLength < 10) throw new TypeError('Truncated GIF logical screen descriptor');
  const width = readUint16LittleEndian(bytes, 6);
  const height = readUint16LittleEndian(bytes, 8);
  if (width === 0 || height === 0) throw new TypeError('GIF dimensions must be positive');
  return Object.freeze({ contentType: 'image/gif', width, height });
}

function inspectJpeg(bytes: Uint8Array): RasterImageInfo {
  let offset = 2;
  while (offset < bytes.byteLength) {
    if (bytes[offset] !== 0xff) throw new TypeError('Invalid JPEG marker prefix');
    while (offset < bytes.byteLength && bytes[offset] === 0xff) offset += 1;
    if (offset >= bytes.byteLength) throw new TypeError('Truncated JPEG marker');
    const marker = bytes[offset]!;
    offset += 1;
    if (marker === 0x00) throw new TypeError('Invalid JPEG marker 0x00 before scan data');
    if (marker === 0xda) throw new TypeError('JPEG reached SOS before SOF');
    if (marker === 0xd9) throw new TypeError('JPEG reached EOI before SOF');
    if (isStandaloneJpegMarker(marker)) continue;
    if (offset + 2 > bytes.byteLength) throw new TypeError('Truncated JPEG segment length');
    const length = readUint16BigEndian(bytes, offset);
    if (length < 2) throw new TypeError('Invalid JPEG segment length');
    const segmentEnd = offset + length;
    if (segmentEnd > bytes.byteLength) throw new TypeError('Truncated JPEG segment');
    if (JPEG_SOF_MARKERS.has(marker)) {
      if (length < 8) throw new TypeError('JPEG SOF segment is too short');
      const height = readUint16BigEndian(bytes, offset + 3);
      const width = readUint16BigEndian(bytes, offset + 5);
      if (width === 0 || height === 0) throw new TypeError('JPEG dimensions must be positive');
      return Object.freeze({ contentType: 'image/jpeg', width, height });
    }
    offset = segmentEnd;
  }
  throw new TypeError('JPEG contains no SOF marker');
}

function startsWith(bytes: Uint8Array, signature: Uint8Array): boolean {
  if (bytes.byteLength < signature.byteLength) return false;
  for (let index = 0; index < signature.byteLength; index += 1) {
    if (bytes[index] !== signature[index]) return false;
  }
  return true;
}

function isRasterImageCandidate(bytes: Uint8Array): boolean {
  return startsWith(bytes, PNG_SIGNATURE)
    || startsWith(bytes, GIF87A_SIGNATURE)
    || startsWith(bytes, GIF89A_SIGNATURE)
    || isTruncatedPrefix(bytes, PNG_SIGNATURE)
    || isTruncatedPrefix(bytes, GIF87A_SIGNATURE)
    || isTruncatedPrefix(bytes, GIF89A_SIGNATURE)
    || bytes[0] === 0xff;
}

function isTruncatedPrefix(bytes: Uint8Array, signature: Uint8Array): boolean {
  if (bytes.byteLength >= signature.byteLength) return false;
  for (let index = 0; index < bytes.byteLength; index += 1) {
    if (bytes[index] !== signature[index]) return false;
  }
  return true;
}

function readUint32BigEndian(bytes: Uint8Array, offset: number): number {
  return (
    bytes[offset]! * 0x1000000 +
    bytes[offset + 1]! * 0x10000 +
    bytes[offset + 2]! * 0x100 +
    bytes[offset + 3]!
  );
}

function readUint16BigEndian(bytes: Uint8Array, offset: number): number {
  return bytes[offset]! * 0x100 + bytes[offset + 1]!;
}

function readUint16LittleEndian(bytes: Uint8Array, offset: number): number {
  return bytes[offset]! + bytes[offset + 1]! * 0x100;
}

function isStandaloneJpegMarker(marker: number): boolean {
  return marker === 0x01 || marker === 0xd8 || (marker >= 0xd0 && marker <= 0xd7);
}

function isBlob(value: unknown): value is Blob {
  return typeof Blob !== 'undefined' && value instanceof Blob;
}

function isReadableStream(value: unknown): value is ReadableStream<unknown> {
  return Boolean(value && typeof (value as { getReader?: unknown }).getReader === 'function');
}

function isAsyncIterable(value: unknown): value is AsyncIterable<unknown> {
  return Boolean(
    value &&
    typeof (value as { [Symbol.asyncIterator]?: unknown })[Symbol.asyncIterator] === 'function'
  );
}

async function readWebStream(
  stream: ReadableStream<unknown>,
  signal?: AbortSignal,
): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  const reader = stream.getReader();
  const cancelOnAbort = (): void => {
    void reader.cancel(signal?.reason).catch(() => undefined);
  };
  signal?.addEventListener('abort', cancelOnAbort, { once: true });
  try {
    while (true) {
      throwIfAborted(signal);
      let result: ReadableStreamReadResult<unknown>;
      try {
        result = await reader.read();
      } catch (error) {
        throwIfAborted(signal);
        throw error;
      }
      throwIfAborted(signal);
      if (result.done) break;
      chunks.push(normalizeByteChunk(result.value));
    }
  } finally {
    signal?.removeEventListener('abort', cancelOnAbort);
    reader.releaseLock();
  }
  return concatenateBytes(chunks);
}

async function readAsyncIterable(
  source: AsyncIterable<unknown>,
  signal?: AbortSignal,
): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  const iterator = source[Symbol.asyncIterator]();
  let completed = false;
  try {
    while (true) {
      throwIfAborted(signal);
      const result = await iterator.next();
      throwIfAborted(signal);
      if (result.done) {
        completed = true;
        break;
      }
      chunks.push(normalizeByteChunk(result.value));
    }
  } finally {
    if (!completed && typeof iterator.return === 'function') await iterator.return();
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
  throw new TypeError(
    'Image streams must yield byte numbers, Uint8Array, ArrayBuffer, or ArrayBufferView chunks',
  );
}

function concatenateBytes(chunks: readonly Uint8Array[]): Uint8Array {
  const length = chunks.reduce((total, chunk) => total + chunk.byteLength, 0);
  const result = new Uint8Array(length);
  let offset = 0;
  for (const chunk of chunks) {
    result.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return result;
}

function throwIfAborted(signal?: AbortSignal): void {
  if (!signal?.aborted) return;
  throw signal.reason instanceof Error
    ? signal.reason
    : new DOMException('Image loading was aborted', 'AbortError');
}
