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
