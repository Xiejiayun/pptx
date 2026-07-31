import { describe, expect, it } from 'vitest';
import {
  inspectRasterImage,
  resolveRasterImageSource,
  type RasterImageByteChunk,
} from './raster-image-source.js';

const PNG_SIGNATURE = [137, 80, 78, 71, 13, 10, 26, 10] as const;
const JPEG_SOF_MARKERS = [
  0xc0, 0xc1, 0xc2, 0xc3,
  0xc5, 0xc6, 0xc7,
  0xc9, 0xca, 0xcb,
  0xcd, 0xce, 0xcf,
] as const;

function pngHeader(width: number, height: number): Uint8Array<ArrayBuffer> {
  const bytes = new Uint8Array(24);
  bytes.set(PNG_SIGNATURE);
  writeUint32(bytes, 8, 13);
  bytes.set([73, 72, 68, 82], 12);
  writeUint32(bytes, 16, width);
  writeUint32(bytes, 20, height);
  return bytes;
}

function gifHeader(
  version: 'GIF87a' | 'GIF89a',
  width: number,
  height: number,
): Uint8Array<ArrayBuffer> {
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
): Uint8Array<ArrayBuffer> {
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

describe('raster image in-memory source resolution', () => {
  it('copies Uint8Array and ArrayBuffer sources before returning', async () => {
    const uint8Source = pngHeader(640, 360);
    const uint8Expected = new Uint8Array(uint8Source);
    const uint8Promise = resolveRasterImageSource(uint8Source);
    uint8Source.fill(0);
    const uint8Resolved = await uint8Promise;
    expect(uint8Resolved).toEqual({
      bytes: uint8Expected,
      info: { contentType: 'image/png', width: 640, height: 360 },
    });
    expect(uint8Resolved.bytes.buffer).not.toBe(uint8Source.buffer);

    const arrayBufferSource = gifHeader('GIF89a', 320, 200).buffer;
    const arrayBufferExpected = new Uint8Array(arrayBufferSource.slice(0));
    const arrayBufferPromise = resolveRasterImageSource(arrayBufferSource);
    new Uint8Array(arrayBufferSource).fill(0);
    const arrayBufferResolved = await arrayBufferPromise;
    expect(arrayBufferResolved).toEqual({
      bytes: arrayBufferExpected,
      info: { contentType: 'image/gif', width: 320, height: 200 },
    });
    expect(arrayBufferResolved.bytes.buffer).not.toBe(arrayBufferSource);
  });

  it('loads Blob and File sources while ignoring metadata hints', async () => {
    const jpeg = jpegWithSof(0xc0, 1920, 1080);
    const blob = new Blob([jpeg], { type: 'image/png' });
    await expect(resolveRasterImageSource(blob)).resolves.toEqual({
      bytes: jpeg,
      info: { contentType: 'image/jpeg', width: 1920, height: 1080 },
    });

    if (typeof File !== 'undefined') {
      const file = new File([jpeg], 'misleading.gif', { type: 'application/octet-stream' });
      await expect(resolveRasterImageSource(file)).resolves.toMatchObject({
        info: { contentType: 'image/jpeg', width: 1920, height: 1080 },
      });
    }
  });

  it('checks abort state around Blob reads and propagates Blob failures', async () => {
    const before = new AbortController();
    const beforeReason = new Error('abort before Blob');
    before.abort(beforeReason);
    await expect(resolveRasterImageSource(new Blob([pngHeader(1, 1)]), before.signal))
      .rejects.toBe(beforeReason);

    const after = new AbortController();
    const afterReason = new Error('abort after Blob');
    class AbortingBlob extends Blob {
      override async arrayBuffer(): Promise<ArrayBuffer> {
        const result = await super.arrayBuffer();
        after.abort(afterReason);
        return result;
      }
    }
    await expect(resolveRasterImageSource(new AbortingBlob([pngHeader(1, 1)]), after.signal))
      .rejects.toBe(afterReason);

    const failure = new Error('Blob read failed');
    class FailingBlob extends Blob {
      override arrayBuffer(): Promise<ArrayBuffer> {
        return Promise.reject(failure);
      }
    }
    await expect(resolveRasterImageSource(new FailingBlob())).rejects.toBe(failure);
  });

  it('normalizes every supported Web stream chunk type and releases the reader', async () => {
    const png = pngHeader(640, 360);
    const paddedView = new Uint8Array(8);
    paddedView.set(png.slice(12, 16), 2);
    const chunks: RasterImageByteChunk[] = [
      png[0]!,
      png.slice(1, 8),
      png.slice(8, 12).buffer,
      new DataView(paddedView.buffer, 2, 4),
      new Uint16Array(png.slice(16).buffer),
    ];
    const stream = new ReadableStream<RasterImageByteChunk>({
      start(controller) {
        for (const chunk of chunks) controller.enqueue(chunk);
        controller.close();
      },
    });

    const resolved = await resolveRasterImageSource(stream);
    expect(resolved.bytes).toEqual(png);
    expect(resolved.info).toEqual({ contentType: 'image/png', width: 640, height: 360 });
    expect(stream.locked).toBe(false);
  });

  it('copies async iterable chunks before requesting the next chunk', async () => {
    const png = pngHeader(640, 360);
    const first = png.slice(0, 12);
    let closed = false;
    const source: AsyncIterable<RasterImageByteChunk> = {
      async *[Symbol.asyncIterator]() {
        try {
          yield first;
          first.fill(0);
          yield png.slice(12);
        } finally {
          closed = true;
        }
      },
    };

    const resolved = await resolveRasterImageSource(source);
    expect(resolved.bytes).toEqual(png);
    expect(closed).toBe(true);
  });

  it('closes async iterators and releases Web readers on invalid chunks', async () => {
    let closed = false;
    const iterable: AsyncIterable<unknown> = {
      async *[Symbol.asyncIterator]() {
        try {
          yield pngHeader(1, 1).slice(0, 8);
          yield 'unsafe';
        } finally {
          closed = true;
        }
      },
    };
    await expect(resolveRasterImageSource(iterable as never)).rejects.toThrow(/streams must yield/i);
    expect(closed).toBe(true);

    const stream = new ReadableStream<unknown>({
      start(controller) {
        controller.enqueue('unsafe');
      },
    });
    await expect(resolveRasterImageSource(stream as never)).rejects.toThrow(/streams must yield/i);
    expect(stream.locked).toBe(false);
  });

  it('propagates read failures while closing the source', async () => {
    const streamFailure = new Error('Web stream failed');
    const stream = new ReadableStream<RasterImageByteChunk>({
      pull() {
        throw streamFailure;
      },
    });
    await expect(resolveRasterImageSource(stream)).rejects.toBe(streamFailure);
    expect(stream.locked).toBe(false);

    const iteratorFailure = new Error('Async iterator failed');
    let returned = false;
    const source: AsyncIterable<RasterImageByteChunk> = {
      [Symbol.asyncIterator]() {
        return {
          next: () => Promise.reject(iteratorFailure),
          return: async () => {
            returned = true;
            return { done: true, value: undefined };
          },
        };
      },
    };
    await expect(resolveRasterImageSource(source)).rejects.toBe(iteratorFailure);
    expect(returned).toBe(true);
  });

  it('cancels a pending Web stream read when aborted', async () => {
    const controller = new AbortController();
    const reason = new Error('abort pending Web read');
    let cancelledWith: unknown;
    const stream = new ReadableStream<RasterImageByteChunk>({
      pull() {},
      cancel(value) {
        cancelledWith = value;
      },
    });
    const loading = resolveRasterImageSource(stream, controller.signal);
    controller.abort(reason);

    await expect(loading).rejects.toBe(reason);
    expect(cancelledWith).toBe(reason);
    expect(stream.locked).toBe(false);
  });

  it('stops and closes an async iterator when aborted between chunks', async () => {
    const controller = new AbortController();
    const reason = new Error('stop raster stream');
    let calls = 0;
    let returned = false;
    const iterator: AsyncIterator<RasterImageByteChunk> = {
      async next() {
        calls += 1;
        if (calls === 1) {
          controller.abort(reason);
          return { done: false, value: pngHeader(1, 1).slice(0, 8) };
        }
        return { done: false, value: pngHeader(1, 1).slice(8) };
      },
      async return() {
        returned = true;
        return { done: true, value: undefined };
      },
    };
    const source: AsyncIterable<RasterImageByteChunk> = {
      [Symbol.asyncIterator]: () => iterator,
    };

    await expect(resolveRasterImageSource(source, controller.signal)).rejects.toBe(reason);
    expect(calls).toBe(1);
    expect(returned).toBe(true);
  });

  it.each([
    { name: 'fractional byte', chunk: 1.5 },
    { name: 'negative byte', chunk: -1 },
    { name: 'large byte', chunk: 256 },
    { name: 'string', chunk: '1' },
    { name: 'plain object', chunk: {} },
  ])('rejects a $name stream chunk', async ({ chunk }) => {
    const source: AsyncIterable<unknown> = {
      async *[Symbol.asyncIterator]() {
        yield chunk;
      },
    };
    await expect(resolveRasterImageSource(source as never)).rejects.toThrow(/streams must yield/i);
  });

  it.each([
    { name: 'empty bytes', source: new Uint8Array() },
    { name: 'empty Blob', source: new Blob() },
    { name: 'plain object', source: {} },
    { name: 'sync iterable', source: [1, 2, 3] },
    { name: 'null', source: null },
  ])('rejects unsupported or invalid $name sources', async ({ source }) => {
    await expect(resolveRasterImageSource(source as never)).rejects.toThrow(TypeError);
  });
});
