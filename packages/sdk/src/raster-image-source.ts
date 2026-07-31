import type { RasterImageContentType } from '@pptx/model';

const PNG_SIGNATURE = Uint8Array.from([137, 80, 78, 71, 13, 10, 26, 10]);
const GIF87A_SIGNATURE = Uint8Array.from([71, 73, 70, 56, 55, 97]);
const GIF89A_SIGNATURE = Uint8Array.from([71, 73, 70, 56, 57, 97]);
const JPEG_SOF_MARKERS = new Set([
  0xc0, 0xc1, 0xc2, 0xc3,
  0xc5, 0xc6, 0xc7,
  0xc9, 0xca, 0xcb,
  0xcd, 0xce, 0xcf,
]);

export interface RasterImageInfo {
  readonly contentType: RasterImageContentType;
  readonly width: number;
  readonly height: number;
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

export async function resolveRasterImageSource(
  source: RasterImageSource,
  signal?: AbortSignal,
): Promise<ResolvedRasterImageSource> {
  throwIfAborted(signal);
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
    throw new TypeError('Unsupported raster image source type');
  }
  throwIfAborted(signal);
  return { bytes, info: inspectRasterImage(bytes) };
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
    'Raster image streams must yield byte numbers, Uint8Array, ArrayBuffer, or ArrayBufferView chunks',
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
    : new DOMException('Raster image loading was aborted', 'AbortError');
}
