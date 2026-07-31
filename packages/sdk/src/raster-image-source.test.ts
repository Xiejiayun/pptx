import { describe, expect, it } from 'vitest';
import { inspectRasterImage } from './raster-image-source.js';

const PNG_SIGNATURE = [137, 80, 78, 71, 13, 10, 26, 10] as const;
const JPEG_SOF_MARKERS = [
  0xc0, 0xc1, 0xc2, 0xc3,
  0xc5, 0xc6, 0xc7,
  0xc9, 0xca, 0xcb,
  0xcd, 0xce, 0xcf,
] as const;

function pngHeader(width: number, height: number): Uint8Array {
  const bytes = new Uint8Array(24);
  bytes.set(PNG_SIGNATURE);
  writeUint32(bytes, 8, 13);
  bytes.set([73, 72, 68, 82], 12);
  writeUint32(bytes, 16, width);
  writeUint32(bytes, 20, height);
  return bytes;
}

function gifHeader(version: 'GIF87a' | 'GIF89a', width: number, height: number): Uint8Array {
  const bytes = new Uint8Array(10);
  bytes.set([...version].map((character) => character.charCodeAt(0)));
  bytes[6] = width & 0xff;
  bytes[7] = width >>> 8;
  bytes[8] = height & 0xff;
  bytes[9] = height >>> 8;
  return bytes;
}

function jpegWithSof(
  marker: number,
  width: number,
  height: number,
  prefix: readonly number[] = [],
): Uint8Array {
  return Uint8Array.from([
    0xff, 0xd8,
    ...prefix,
    0xff, marker,
    0x00, 0x08,
    0x08,
    height >>> 8, height & 0xff,
    width >>> 8, width & 0xff,
    0x01,
  ]);
}

function writeUint32(bytes: Uint8Array, offset: number, value: number): void {
  bytes[offset] = value >>> 24;
  bytes[offset + 1] = value >>> 16;
  bytes[offset + 2] = value >>> 8;
  bytes[offset + 3] = value;
}

describe('raster image inspection', () => {
  it('inspects PNG and GIF signatures without changing the input', () => {
    const png = pngHeader(640, 360);
    const before = new Uint8Array(png);
    const info = inspectRasterImage(png);

    expect(info).toEqual({ contentType: 'image/png', width: 640, height: 360 });
    expect(Object.isFrozen(info)).toBe(true);
    expect(inspectRasterImage(png)).toEqual(info);
    expect(png).toEqual(before);
    expect(inspectRasterImage(gifHeader('GIF87a', 320, 200))).toEqual({
      contentType: 'image/gif',
      width: 320,
      height: 200,
    });
    expect(inspectRasterImage(gifHeader('GIF89a', 65_535, 1))).toEqual({
      contentType: 'image/gif',
      width: 65_535,
      height: 1,
    });
  });

  it('inspects every JPEG SOF family that carries sample dimensions', () => {
    for (const marker of JPEG_SOF_MARKERS) {
      expect(inspectRasterImage(jpegWithSof(marker, 1920, 1080)), marker.toString(16)).toEqual({
        contentType: 'image/jpeg',
        width: 1920,
        height: 1080,
      });
    }
  });

  it('safely traverses JPEG segments, fill bytes, and standalone markers', () => {
    const prefix = [
      0xff, 0xe1, 0x00, 0x04, 0x11, 0x22,
      0xff, 0xc4, 0x00, 0x02,
      0xff, 0xc8, 0x00, 0x02,
      0xff, 0xcc, 0x00, 0x02,
      0xff, 0xff, 0x01,
      0xff, 0xd0,
    ];
    expect(inspectRasterImage(jpegWithSof(0xc2, 800, 600, prefix))).toEqual({
      contentType: 'image/jpeg',
      width: 800,
      height: 600,
    });
  });

  it.each([
    { name: 'non-byte input', bytes: [1, 2, 3] as unknown, error: /Uint8Array/ },
    { name: 'empty input', bytes: new Uint8Array(), error: /empty/i },
    { name: 'unknown signature', bytes: Uint8Array.of(1, 2, 3), error: /unsupported.*signature/i },
    { name: 'truncated PNG signature', bytes: Uint8Array.from(PNG_SIGNATURE.slice(0, 7)), error: /truncated PNG signature/i },
    { name: 'truncated PNG IHDR', bytes: Uint8Array.from(PNG_SIGNATURE), error: /truncated PNG IHDR/i },
    { name: 'wrong PNG first length', bytes: (() => { const value = pngHeader(1, 1); value[11] = 12; return value; })(), error: /PNG.*IHDR length/i },
    { name: 'wrong PNG first chunk', bytes: (() => { const value = pngHeader(1, 1); value[12] = 88; return value; })(), error: /PNG.*first chunk/i },
    { name: 'zero PNG width', bytes: pngHeader(0, 1), error: /PNG.*dimensions/i },
    { name: 'zero PNG height', bytes: pngHeader(1, 0), error: /PNG.*dimensions/i },
    { name: 'truncated GIF signature', bytes: Uint8Array.from([71, 73, 70, 56, 57]), error: /truncated GIF signature/i },
    { name: 'truncated GIF descriptor', bytes: gifHeader('GIF89a', 1, 1).slice(0, 9), error: /truncated GIF.*descriptor/i },
    { name: 'zero GIF width', bytes: gifHeader('GIF89a', 0, 1), error: /GIF.*dimensions/i },
    { name: 'zero GIF height', bytes: gifHeader('GIF89a', 1, 0), error: /GIF.*dimensions/i },
    { name: 'truncated JPEG signature', bytes: Uint8Array.of(0xff), error: /truncated JPEG signature/i },
    { name: 'missing JPEG marker', bytes: Uint8Array.of(0xff, 0xd8, 0x01), error: /JPEG marker/i },
    { name: 'truncated JPEG marker', bytes: Uint8Array.of(0xff, 0xd8, 0xff), error: /truncated JPEG marker/i },
    { name: 'stuffed JPEG marker', bytes: Uint8Array.of(0xff, 0xd8, 0xff, 0x00), error: /invalid JPEG marker/i },
    { name: 'truncated JPEG length', bytes: Uint8Array.of(0xff, 0xd8, 0xff, 0xe0, 0x00), error: /truncated JPEG segment length/i },
    { name: 'short JPEG length', bytes: Uint8Array.of(0xff, 0xd8, 0xff, 0xe0, 0x00, 0x01), error: /invalid JPEG segment length/i },
    { name: 'truncated JPEG segment', bytes: Uint8Array.of(0xff, 0xd8, 0xff, 0xe0, 0x00, 0x04, 0x01), error: /truncated JPEG segment/i },
    { name: 'short JPEG SOF', bytes: Uint8Array.of(0xff, 0xd8, 0xff, 0xc0, 0x00, 0x07, 8, 0, 1, 0, 1), error: /JPEG SOF.*short/i },
    { name: 'zero JPEG width', bytes: jpegWithSof(0xc0, 0, 1), error: /JPEG.*dimensions/i },
    { name: 'zero JPEG height', bytes: jpegWithSof(0xc0, 1, 0), error: /JPEG.*dimensions/i },
    { name: 'JPEG SOS before SOF', bytes: Uint8Array.of(0xff, 0xd8, 0xff, 0xda), error: /JPEG.*SOS.*before SOF/i },
    { name: 'JPEG EOI before SOF', bytes: Uint8Array.of(0xff, 0xd8, 0xff, 0xd9), error: /JPEG.*EOI.*before SOF/i },
    { name: 'JPEG without SOF', bytes: Uint8Array.of(0xff, 0xd8, 0xff, 0xe0, 0x00, 0x02), error: /JPEG.*no SOF/i },
  ])('rejects $name', ({ bytes, error }) => {
    expect(() => inspectRasterImage(bytes as Uint8Array)).toThrow(TypeError);
    expect(() => inspectRasterImage(bytes as Uint8Array)).toThrow(error);
  });
});
