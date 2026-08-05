import { createServer } from 'node:http';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, relative } from 'node:path';
import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  assertRasterImageContentType,
  inspectRasterImage,
  normalizeAddImageSourceOptions,
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

function dataUri(contentType: 'image/png' | 'image/jpeg' | 'image/gif', bytes: Uint8Array): string {
  let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return `data:${contentType};base64,${btoa(binary)}`;
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

describe('raster image data URI resolution', () => {
  it('decodes canonical PNG, JPEG, and GIF data URIs', async () => {
    const cases = [
      { contentType: 'image/png' as const, bytes: pngHeader(640, 360), width: 640, height: 360 },
      { contentType: 'image/jpeg' as const, bytes: jpegWithSof(0xc0, 1920, 1080), width: 1920, height: 1080 },
      { contentType: 'image/gif' as const, bytes: gifHeader('GIF89a', 320, 200), width: 320, height: 200 },
    ];
    for (const fixture of cases) {
      await expect(resolveRasterImageSource(dataUri(fixture.contentType, fixture.bytes))).resolves.toEqual({
        bytes: fixture.bytes,
        info: {
          contentType: fixture.contentType,
          width: fixture.width,
          height: fixture.height,
        },
        assertedContentType: fixture.contentType,
      });
    }
  });

  it('accepts case-insensitive data URI media type and base64 tokens', async () => {
    const source = dataUri('image/png', pngHeader(3, 2))
      .replace('data:image/png;base64,', 'DATA:IMAGE/PNG;BASE64,');
    await expect(resolveRasterImageSource(source)).resolves.toMatchObject({
      info: { contentType: 'image/png', width: 3, height: 2 },
      assertedContentType: 'image/png',
    });
  });

  it.each([
    { name: 'missing media type', source: 'data:;base64,AAAA' },
    { name: 'empty media type', source: 'data:base64,AAAA' },
    { name: 'JPG alias', source: 'data:image/jpg;base64,AAAA' },
    { name: 'SVG media type', source: 'data:image/svg+xml;base64,AAAA' },
    { name: 'missing base64 token', source: 'data:image/png,AAAA' },
    { name: 'extra parameter', source: 'data:image/png;charset=utf-8;base64,AAAA' },
    { name: 'raw percent encoding', source: 'data:image/png;base64,%89PNG' },
    { name: 'empty payload', source: 'data:image/png;base64,' },
    { name: 'payload whitespace', source: 'data:image/png;base64,AA==\n' },
    { name: 'URL-safe alphabet', source: 'data:image/png;base64,____' },
    { name: 'short group', source: 'data:image/png;base64,AAA' },
    { name: 'too much padding', source: 'data:image/png;base64,A===' },
    { name: 'unexpected padding', source: 'data:image/png;base64,AA=A' },
    { name: 'extra padding', source: 'data:image/png;base64,AAAA=' },
    { name: 'non-zero double-padding bits', source: 'data:image/png;base64,AB==' },
    { name: 'non-zero single-padding bits', source: 'data:image/png;base64,AAB=' },
    { name: 'trailing data', source: 'data:image/png;base64,AAAA,AAAA' },
  ])('rejects $name', async ({ source }) => {
    await expect(resolveRasterImageSource(source)).rejects.toThrow(TypeError);
    await expect(resolveRasterImageSource(source)).rejects.toThrow(/data URI|signature/i);
  });

  it('rejects a declared MIME that disagrees with the detected signature', async () => {
    const payload = dataUri('image/png', pngHeader(8, 6)).split(',')[1]!;
    await expect(resolveRasterImageSource(`data:image/gif;base64,${payload}`))
      .rejects.toThrow(/declares image\/gif.*signature.*image\/png/i);
  });

  it('does not expose a rejected data URI payload in errors', async () => {
    const payload = 'UNIQUESECRETINVALIDPAYLOAD';
    try {
      await resolveRasterImageSource(`data:image/png;base64,${payload}`);
      throw new Error('Expected data URI rejection');
    } catch (error) {
      expect(error).toBeInstanceOf(TypeError);
      expect((error as Error).message).not.toContain(payload);
    }
  });
});

describe('raster image path and URL resolution', () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('loads relative and absolute Node paths using signature detection', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'pptx-raster-source-'));
    try {
      const fixtures = [
        { name: 'misleading.bin', bytes: pngHeader(640, 360), contentType: 'image/png' },
        { name: 'photo.data', bytes: jpegWithSof(0xc0, 1920, 1080), contentType: 'image/jpeg' },
        { name: 'animation.unknown', bytes: gifHeader('GIF89a', 320, 200), contentType: 'image/gif' },
      ] as const;
      for (const fixture of fixtures) {
        const path = join(directory, fixture.name);
        await writeFile(path, fixture.bytes);
        const absolute = await resolveRasterImageSource(path);
        expect(absolute.bytes).toEqual(fixture.bytes);
        expect(absolute.info.contentType).toBe(fixture.contentType);
        const relativePath = relative(process.cwd(), path);
        expect((await resolveRasterImageSource(relativePath)).bytes).toEqual(fixture.bytes);
      }
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  });

  it('propagates local path errors, parser failures, and aborts', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'pptx-raster-source-error-'));
    try {
      const truncated = join(directory, 'truncated.png');
      await writeFile(truncated, Uint8Array.from(PNG_SIGNATURE));
      await expect(resolveRasterImageSource(truncated)).rejects.toThrow(/truncated PNG IHDR/i);
      await expect(resolveRasterImageSource(join(directory, 'missing.png'))).rejects.toMatchObject({
        code: 'ENOENT',
      });
      await expect(resolveRasterImageSource(directory)).rejects.toBeInstanceOf(Error);

      const controller = new AbortController();
      const reason = new Error('abort local raster path');
      controller.abort(reason);
      await expect(resolveRasterImageSource(truncated, controller.signal)).rejects.toBe(reason);
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  });

  it('treats Windows drive strings as paths and rejects unsupported URL schemes', async () => {
    let windowsPathError: unknown;
    try {
      await resolveRasterImageSource('C:\\missing\\image.png');
    } catch (error) {
      windowsPathError = error;
    }
    expect(windowsPathError).toBeInstanceOf(Error);
    expect((windowsPathError as Error).message).not.toMatch(/unsupported.*scheme/i);
    for (const source of ['', 'ftp://example.com/image.png', 'file:///tmp/image.png', 'mailto:image@example.com']) {
      await expect(resolveRasterImageSource(source)).rejects.toThrow(TypeError);
      if (source) await expect(resolveRasterImageSource(source)).rejects.toThrow(/unsupported.*scheme/i);
    }
  });

  it('downloads HTTP images, follows redirects, ignores response MIME, and supports abort', async () => {
    const png = pngHeader(800, 450);
    let slowRequestedResolve!: () => void;
    const slowRequested = new Promise<void>((resolve) => {
      slowRequestedResolve = resolve;
    });
    const server = createServer((request, response) => {
      const url = new URL(request.url ?? '/', 'http://127.0.0.1');
      if (url.pathname === '/redirect') {
        response.writeHead(302, { location: '/image.png?from=redirect' });
        response.end();
        return;
      }
      if (url.pathname === '/image.png') {
        response.writeHead(200, { 'content-type': 'application/octet-stream' });
        response.end(png);
        return;
      }
      if (url.pathname === '/truncated') {
        response.writeHead(200, { 'content-type': 'image/png' });
        response.end(png.slice(0, 8));
        return;
      }
      if (url.pathname === '/slow') {
        slowRequestedResolve();
        return;
      }
      response.writeHead(404, { 'content-type': 'image/png' });
      response.end('missing');
    });
    await new Promise<void>((resolve, reject) => {
      server.once('error', reject);
      server.listen(0, '127.0.0.1', resolve);
    });
    const address = server.address();
    if (!address || typeof address === 'string') throw new Error('HTTP test server has no port');
    const baseUrl = `http://127.0.0.1:${address.port}`;
    try {
      await expect(resolveRasterImageSource(`${baseUrl}/image.png?cache=1`)).resolves.toEqual({
        bytes: png,
        info: { contentType: 'image/png', width: 800, height: 450 },
      });
      await expect(resolveRasterImageSource(`${baseUrl}/redirect`)).resolves.toMatchObject({
        info: { contentType: 'image/png', width: 800, height: 450 },
      });
      await expect(resolveRasterImageSource(`${baseUrl}/missing`)).rejects.toThrow(/HTTP 404/i);
      await expect(resolveRasterImageSource(`${baseUrl}/truncated`)).rejects.toThrow(/truncated PNG IHDR/i);

      const controller = new AbortController();
      const reason = new Error('abort HTTP raster source');
      const loading = resolveRasterImageSource(`${baseUrl}/slow`, controller.signal);
      await slowRequested;
      controller.abort(reason);
      await expect(loading).rejects.toBe(reason);
    } finally {
      server.closeAllConnections();
      await new Promise<void>((resolve, reject) => {
        server.close((error) => error ? reject(error) : resolve());
      });
    }
    await expect(resolveRasterImageSource(`${baseUrl}/image.png`)).rejects.toBeInstanceOf(Error);
  });

  it('uses Fetch for browser-relative URLs and preserves the request string', async () => {
    const png = pngHeader(16, 9);
    const fetchMock = vi.fn(async (input: RequestInfo | URL) => {
      if (String(input).includes('missing')) return new Response(null, { status: 404 });
      return new Response(png, {
        status: 200,
        headers: { 'content-type': 'text/plain' },
      });
    });
    vi.stubGlobal('fetch', fetchMock);
    vi.stubGlobal('process', undefined);

    await expect(resolveRasterImageSource('./assets/image%20one.png?version=1')).resolves.toMatchObject({
      info: { contentType: 'image/png', width: 16, height: 9 },
    });
    await expect(resolveRasterImageSource('/assets/root.png')).resolves.toMatchObject({
      info: { contentType: 'image/png', width: 16, height: 9 },
    });
    await expect(resolveRasterImageSource('/assets/missing.png')).rejects.toThrow(/HTTP 404/i);
    expect(fetchMock.mock.calls.map(([input]) => input)).toEqual([
      './assets/image%20one.png?version=1',
      '/assets/root.png',
      '/assets/missing.png',
    ]);
  });
});

describe('raster image document options', () => {
  it('detaches image fields, optional content type, and AbortSignal', () => {
    const controller = new AbortController();
    const fallback = pngHeader(1, 1);
    const shadow = {
      kind: 'outer' as const,
      color: { kind: 'srgb' as const, value: '123456' },
      opacity: 0.5,
      rotateWithShape: true,
    };
    const options = {
      contentType: 'image/svg+xml' as const,
      fallback,
      signal: controller.signal,
      name: 'Source image',
      altText: '',
      x: -1,
      y: 2,
      width: 3,
      height: 4,
      rotation: 5,
      flipHorizontal: true,
      flipVertical: false,
      rounding: true,
      shadow,
      transparency: 25,
    };
    const normalized = normalizeAddImageSourceOptions(options);
    options.name = 'Changed';
    shadow.color.value = 'FFFFFF';
    fallback.fill(0);

    expect(normalized.contentType).toBe('image/svg+xml');
    expect(normalized.fallback).toEqual(pngHeader(1, 1));
    expect((normalized.fallback as Uint8Array).buffer).not.toBe(fallback.buffer);
    expect(normalized.signal).toBe(controller.signal);
    expect(normalized.imageOptions).toEqual({
      name: 'Source image',
      altText: '',
      x: -1,
      y: 2,
      width: 3,
      height: 4,
      rotation: 5,
      flipHorizontal: true,
      flipVertical: false,
      rounding: true,
      shadow: {
        kind: 'outer',
        color: { kind: 'srgb', value: '123456' },
        opacity: 0.5,
        rotateWithShape: true,
      },
      transparency: 25,
    });
    expect(Object.getPrototypeOf(normalized.imageOptions)).toBeNull();
    expect(Object.isFrozen(normalized.imageOptions)).toBe(true);
    expect(Object.isFrozen(normalized.imageOptions.shadow)).toBe(true);
    expect(Object.isFrozen(normalized.imageOptions.shadow?.color)).toBe(true);
    expect(Object.isFrozen(normalized)).toBe(true);
  });

  it('preserves detached percentage coordinate source options', () => {
    const options = {
      x: '10%',
      y: '20%',
      width: '30%',
      height: '40%',
    };
    const normalized = normalizeAddImageSourceOptions(options);
    options.x = '90%';

    expect(normalized.imageOptions).toEqual({
      x: '10%',
      y: '20%',
      width: '30%',
      height: '40%',
    });
    expect(Object.getPrototypeOf(normalized.imageOptions)).toBeNull();
    expect(Object.isFrozen(normalized.imageOptions)).toBe(true);
  });

  it('detaches and deeply freezes explicit sizing while keeping placement separate', () => {
    const source = { x: 400, y: 225, width: 800, height: 450 };
    const sizing = {
      type: 'crop' as const,
      width: 3,
      height: 2,
      source,
    };
    const normalized = normalizeAddImageSourceOptions({
      x: 10,
      y: 20,
      rotation: 30,
      sizing,
    });
    sizing.width = 999;
    source.x = 999;

    expect(normalized.imageOptions).toEqual({ x: 10, y: 20, rotation: 30 });
    expect(normalized.sizing).toEqual({
      type: 'crop',
      width: 3,
      height: 2,
      source: { x: 400, y: 225, width: 800, height: 450 },
    });
    expect(Object.isFrozen(normalized.sizing)).toBe(true);
    if (normalized.sizing?.type !== 'crop') throw new Error('Expected normalized crop sizing');
    expect(Object.isFrozen(normalized.sizing.source)).toBe(true);
  });

  it('accepts omitted, explicit undefined, and null-prototype options', () => {
    expect(normalizeAddImageSourceOptions({})).toMatchObject({ imageOptions: {} });
    expect(normalizeAddImageSourceOptions({ contentType: undefined, signal: undefined }))
      .toMatchObject({ imageOptions: {} });
    const options = Object.assign(Object.create(null) as Record<string, unknown>, {
      contentType: 'image/gif',
      name: 'Null prototype',
    });
    expect(normalizeAddImageSourceOptions(options)).toMatchObject({
      contentType: 'image/gif',
      imageOptions: { name: 'Null prototype' },
    });
  });

  it('rejects unsafe options without invoking accessors', () => {
    let reads = 0;
    const accessor = Object.defineProperty({}, 'name', {
      enumerable: true,
      get() {
        reads += 1;
        return 'unsafe';
      },
    });
    const symbol = { [Symbol('unsafe')]: true };
    class Options {}
    const invalid: unknown[] = [
      null,
      [],
      new Options(),
      Object.create({ contentType: 'image/png' }),
      accessor,
      symbol,
      { unknown: true },
      { contentType: 'image/svg' },
      { fallback: {} },
      { fallback: '' },
      { signal: {} },
      { sizing: { type: 'cover', width: 1, height: 1 }, width: 1 },
      { sizing: { type: 'contain', width: 1, height: 1 }, height: undefined },
      { sizing: { type: 'crop', width: 1, height: 1, source: null } },
      { sourceRectangle: { left: 0, top: 0, right: 0, bottom: 0 } },
    ];
    for (const value of invalid) {
      expect(() => normalizeAddImageSourceOptions(value)).toThrow(TypeError);
    }
    expect(reads).toBe(0);
  });

  it('checks explicit and source-declared MIME assertions', () => {
    const resolved = {
      bytes: pngHeader(1, 1),
      info: { contentType: 'image/png' as const, width: 1, height: 1 },
      assertedContentType: 'image/png' as const,
    };
    expect(() => assertRasterImageContentType(undefined, resolved)).not.toThrow();
    expect(() => assertRasterImageContentType('image/png', resolved)).not.toThrow();
    expect(() => assertRasterImageContentType('image/gif', resolved))
      .toThrow(/expected image\/gif.*signature.*image\/png/i);
    expect(() => assertRasterImageContentType(undefined, {
      ...resolved,
      assertedContentType: 'image/jpeg',
    })).toThrow(/declares image\/jpeg.*signature.*image\/png/i);
  });
});
