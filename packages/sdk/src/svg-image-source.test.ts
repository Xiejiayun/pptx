import { createServer } from 'node:http';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, relative } from 'node:path';
import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  assertImageContentType,
  inspectImage,
  resolveImageSource,
  type ImageByteChunk,
  type ImageByteStream,
  type ImageContentType,
  type ImageInfo,
  type ImageSource,
  type ResolvedImageSource,
  inspectSvgImage,
} from './index.js';
import { resolveRasterImageSource } from './raster-image-source.js';

const SVG_NAMESPACE = 'http://www.w3.org/2000/svg';
const PNG_SIGNATURE = [137, 80, 78, 71, 13, 10, 26, 10] as const;

function svgBytes(
  body = `<svg xmlns="${SVG_NAMESPACE}" viewBox="0 0 640 360"/>`,
): Uint8Array<ArrayBuffer> {
  return new TextEncoder().encode(body);
}

function pngHeader(width: number, height: number): Uint8Array<ArrayBuffer> {
  const bytes = new Uint8Array(24);
  bytes.set(PNG_SIGNATURE);
  bytes.set([0, 0, 0, 13, 73, 72, 68, 82], 8);
  bytes[16] = width >>> 24;
  bytes[17] = width >>> 16;
  bytes[18] = width >>> 8;
  bytes[19] = width;
  bytes[20] = height >>> 24;
  bytes[21] = height >>> 16;
  bytes[22] = height >>> 8;
  bytes[23] = height;
  return bytes;
}

function dataUri(contentType: string, bytes: Uint8Array): string {
  let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return `data:${contentType};base64,${btoa(binary)}`;
}

describe('SVG image inspection', () => {
  it.each([
    {
      name: 'viewBox',
      source: `<svg xmlns="${SVG_NAMESPACE}" viewBox="-10.5, 20, 640, 360"/>`,
      width: 640,
      height: 360,
    },
    {
      name: 'unitless dimensions',
      source: `<svg xmlns="${SVG_NAMESPACE}" width="320" height="200"/>`,
      width: 320,
      height: 200,
    },
    {
      name: 'CSS pixels',
      source: `<svg xmlns="${SVG_NAMESPACE}" width="320px" height="200PX"/>`,
      width: 320,
      height: 200,
    },
    {
      name: 'inches',
      source: `<svg xmlns="${SVG_NAMESPACE}" width="2in" height="1in"/>`,
      width: 192,
      height: 96,
    },
    {
      name: 'centimeters',
      source: `<svg xmlns="${SVG_NAMESPACE}" width="2.54cm" height="1.27cm"/>`,
      width: 96,
      height: 48,
    },
    {
      name: 'millimeters',
      source: `<svg xmlns="${SVG_NAMESPACE}" width="25.4mm" height="12.7mm"/>`,
      width: 96,
      height: 48,
    },
    {
      name: 'points',
      source: `<svg xmlns="${SVG_NAMESPACE}" width="72pt" height="36pt"/>`,
      width: 96,
      height: 48,
    },
    {
      name: 'picas',
      source: `<svg xmlns="${SVG_NAMESPACE}" width="6pc" height="3pc"/>`,
      width: 96,
      height: 48,
    },
    {
      name: 'quarter millimeters',
      source: `<svg xmlns="${SVG_NAMESPACE}" width="101.6Q" height="50.8q"/>`,
      width: 96,
      height: 48,
    },
    {
      name: 'width derived through viewBox',
      source: `<svg xmlns="${SVG_NAMESPACE}" width="200" viewBox="0 0 4 2"/>`,
      width: 200,
      height: 100,
    },
    {
      name: 'height derived through viewBox',
      source: `<svg xmlns="${SVG_NAMESPACE}" height="90" viewBox="0 0 4 2"/>`,
      width: 180,
      height: 90,
    },
    {
      name: 'explicit dimensions over viewBox magnitude',
      source: `<svg xmlns="${SVG_NAMESPACE}" width="300" height="100" viewBox="0 0 6 4"/>`,
      width: 300,
      height: 100,
    },
    {
      name: 'default replaced-element size',
      source: `<svg xmlns="${SVG_NAMESPACE}" width="100%" height="auto"/>`,
      width: 300,
      height: 150,
    },
  ])('inspects $name', ({ source, width, height }) => {
    const bytes = svgBytes(source);
    const before = new Uint8Array(bytes);
    const info = inspectSvgImage(bytes);

    expect(info.contentType).toBe('image/svg+xml');
    expect(info.width).toBeCloseTo(width, 12);
    expect(info.height).toBeCloseTo(height, 12);
    expect(Object.isFrozen(info)).toBe(true);
    expect(bytes).toEqual(before);
  });

  it('accepts an XML declaration, comments, and a prefixed SVG root', () => {
    const info = inspectSvgImage(svgBytes(
      `<?xml version="1.0" encoding="UTF-8"?>\n<!-- before -->`
        + `<vector:svg xmlns:vector="${SVG_NAMESPACE}" width="3e2" height="1.5e2">`
        + '<!-- inside --></vector:svg><!-- after -->',
    ));
    expect(info).toEqual({ contentType: 'image/svg+xml', width: 300, height: 150 });
  });

  it.each([
    { name: 'non-byte input', value: '<svg/>' },
    { name: 'empty input', value: new Uint8Array() },
    { name: 'invalid UTF-8', value: Uint8Array.of(0xc3, 0x28) },
    { name: 'truncated XML', value: svgBytes(`<svg xmlns="${SVG_NAMESPACE}">`) },
    { name: 'DTD', value: svgBytes(`<!DOCTYPE svg><svg xmlns="${SVG_NAMESPACE}"/>`) },
    { name: 'entity declaration', value: svgBytes(`<!ENTITY x "unsafe"><svg xmlns="${SVG_NAMESPACE}"/>`) },
    { name: 'multiple roots', value: svgBytes(`<svg xmlns="${SVG_NAMESPACE}"/><svg xmlns="${SVG_NAMESPACE}"/>`) },
    { name: 'HTML root', value: svgBytes(`<html xmlns="${SVG_NAMESPACE}"/>`) },
    { name: 'missing namespace', value: svgBytes('<svg/>') },
    { name: 'wrong namespace', value: svgBytes('<svg xmlns="urn:not-svg"/>') },
    { name: 'nested SVG only', value: svgBytes(`<root><svg xmlns="${SVG_NAMESPACE}"/></root>`) },
    { name: 'text outside root', value: svgBytes(`unsafe<svg xmlns="${SVG_NAMESPACE}"/>`) },
    { name: 'zero viewBox width', value: svgBytes(`<svg xmlns="${SVG_NAMESPACE}" viewBox="0 0 0 10"/>`) },
    { name: 'negative viewBox height', value: svgBytes(`<svg xmlns="${SVG_NAMESPACE}" viewBox="0 0 10 -1"/>`) },
    { name: 'malformed viewBox', value: svgBytes(`<svg xmlns="${SVG_NAMESPACE}" viewBox="0 0 10 20 trailing"/>`) },
  ])('rejects $name', ({ value }) => {
    expect(() => inspectSvgImage(value as never)).toThrow(TypeError);
  });

  it('does not include rejected SVG payloads in errors', () => {
    const secret = 'UNIQUE_SVG_SECRET_PAYLOAD';
    try {
      inspectSvgImage(svgBytes(`<${secret}><svg xmlns="${SVG_NAMESPACE}"/></${secret}>`));
      throw new Error('Expected SVG rejection');
    } catch (error) {
      expect(error).toBeInstanceOf(TypeError);
      expect((error as Error).message).not.toContain(secret);
    }
  });
});

describe('generic image inspection and resolution', () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('dispatches raster signatures and SVG XML without format guessing', () => {
    expect(inspectImage(pngHeader(16, 9))).toEqual({
      contentType: 'image/png',
      width: 16,
      height: 9,
    });
    expect(inspectImage(svgBytes())).toEqual({
      contentType: 'image/svg+xml',
      width: 640,
      height: 360,
    });
    expect(() => inspectImage(Uint8Array.from(PNG_SIGNATURE)))
      .toThrow(/truncated PNG IHDR/i);
    expect(() => inspectImage(svgBytes('<html/>'))).toThrow(TypeError);
  });

  it('copies all portable memory source forms and keeps raster resolution strict', async () => {
    const svg = svgBytes();
    const expected = new Uint8Array(svg);
    const typedSource: ImageSource = svg;
    const uint8Promise: Promise<ResolvedImageSource> = resolveImageSource(typedSource);
    svg.fill(0);
    const uint8 = await uint8Promise;
    expect(uint8).toEqual({
      bytes: expected,
      info: { contentType: 'image/svg+xml', width: 640, height: 360 },
    });
    const info: ImageInfo = uint8.info;
    const contentType: ImageContentType = info.contentType;
    expect(contentType).toBe('image/svg+xml');

    const buffer = expected.buffer.slice(0);
    const bufferPromise = resolveImageSource(buffer);
    new Uint8Array(buffer).fill(0);
    expect((await bufferPromise).bytes).toEqual(expected);

    const blob = new Blob([expected], { type: 'image/png' });
    await expect(resolveImageSource(blob)).resolves.toMatchObject({
      info: { contentType: 'image/svg+xml', width: 640, height: 360 },
    });
    if (typeof File !== 'undefined') {
      const file = new File([expected], 'misleading.png', { type: 'application/octet-stream' });
      await expect(resolveImageSource(file)).resolves.toMatchObject({
        info: { contentType: 'image/svg+xml' },
      });
    }

    await expect(resolveRasterImageSource(expected)).rejects.toThrow(/raster image signature/i);
    await expect(resolveRasterImageSource(dataUri('image/svg+xml', expected)))
      .rejects.toThrow(/data URI|raster image signature/i);
  });

  it('combines Web streams and async iterables with every byte chunk type', async () => {
    const svg = svgBytes();
    const padded = new Uint8Array(svg.length + 4);
    padded.set(svg, 2);
    const stream = new ReadableStream<ImageByteChunk>({
      start(controller) {
        controller.enqueue(svg[0]!);
        controller.enqueue(svg.slice(1, 12));
        controller.enqueue(svg.slice(12, 24).buffer);
        controller.enqueue(new DataView(padded.buffer, 26, svg.length - 24));
        controller.close();
      },
    });
    const typedStream: ImageByteStream = stream;
    expect(typedStream).toBe(stream);
    await expect(resolveImageSource(stream)).resolves.toMatchObject({
      bytes: svg,
      info: { contentType: 'image/svg+xml' },
    });
    expect(stream.locked).toBe(false);

    let closed = false;
    const iterable: AsyncIterable<ImageByteChunk> = {
      async *[Symbol.asyncIterator]() {
        try {
          yield svg.slice(0, 10);
          yield svg.slice(10);
        } finally {
          closed = true;
        }
      },
    };
    await expect(resolveImageSource(iterable)).resolves.toMatchObject({ bytes: svg });
    expect(closed).toBe(true);
  });

  it('accepts only canonical SVG data URIs and verifies their declaration', async () => {
    const svg = svgBytes();
    await expect(resolveImageSource(dataUri('image/svg+xml', svg))).resolves.toEqual({
      bytes: svg,
      info: { contentType: 'image/svg+xml', width: 640, height: 360 },
      assertedContentType: 'image/svg+xml',
    });
    const pngPayload = dataUri('image/png', pngHeader(1, 1)).split(',')[1]!;
    await expect(resolveImageSource(`data:image/svg+xml;base64,${pngPayload}`))
      .rejects.toThrow(/declares image\/svg\+xml.*image\/png/i);

    const invalid = [
      'data:image/svg;base64,AAAA',
      'data:image/svg+xml,AAAA',
      'data:image/svg+xml;charset=utf-8;base64,AAAA',
      'data:image/svg+xml;base64,',
      'data:image/svg+xml;base64,PHN2Zy8%2B',
      'data:image/svg+xml;base64,PHN2Zy8+\n',
      'data:image/svg+xml;base64,PHN2Zy8_',
      'data:image/svg+xml;base64,AAA',
    ];
    for (const source of invalid) {
      await expect(resolveImageSource(source)).rejects.toThrow(TypeError);
    }
  });

  it('checks explicit and source-declared generic content types', () => {
    const svg = svgBytes();
    const resolved = {
      bytes: svg,
      info: { contentType: 'image/svg+xml' as const, width: 640, height: 360 },
      assertedContentType: 'image/svg+xml' as const,
    };
    expect(() => assertImageContentType(undefined, resolved)).not.toThrow();
    expect(() => assertImageContentType('image/svg+xml', resolved)).not.toThrow();
    expect(() => assertImageContentType('image/png', resolved))
      .toThrow(/expected image\/png.*image\/svg\+xml/i);
    expect(() => assertImageContentType(undefined, {
      ...resolved,
      assertedContentType: 'image/png',
    })).toThrow(/declares image\/png.*image\/svg\+xml/i);
  });

  it('loads misleading local paths and detaches their bytes', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'pptx-svg-source-'));
    try {
      const absolutePath = join(directory, 'misleading.png');
      const svg = svgBytes();
      await writeFile(absolutePath, svg);
      await expect(resolveImageSource(absolutePath)).resolves.toMatchObject({
        bytes: svg,
        info: { contentType: 'image/svg+xml', width: 640, height: 360 },
      });
      await expect(resolveImageSource(relative(process.cwd(), absolutePath))).resolves
        .toMatchObject({ info: { contentType: 'image/svg+xml' } });
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  });

  it('downloads SVG through redirects and query URLs while ignoring response MIME', async () => {
    const svg = svgBytes();
    let slowRequestedResolve!: () => void;
    const slowRequested = new Promise<void>((resolve) => {
      slowRequestedResolve = resolve;
    });
    const server = createServer((request, response) => {
      const url = new URL(request.url ?? '/', 'http://127.0.0.1');
      if (url.pathname === '/redirect') {
        response.writeHead(302, { location: '/vector.bin?redirected=1' });
        response.end();
      } else if (url.pathname === '/vector.bin') {
        response.writeHead(200, { 'content-type': 'image/png' });
        response.end(svg);
      } else if (url.pathname === '/malformed') {
        response.writeHead(200, { 'content-type': 'image/svg+xml' });
        response.end('<svg');
      } else if (url.pathname === '/slow') {
        slowRequestedResolve();
      } else {
        response.writeHead(404);
        response.end();
      }
    });
    await new Promise<void>((resolve, reject) => {
      server.once('error', reject);
      server.listen(0, '127.0.0.1', resolve);
    });
    const address = server.address();
    if (!address || typeof address === 'string') throw new Error('HTTP test server has no port');
    const baseUrl = `http://127.0.0.1:${address.port}`;
    try {
      await expect(resolveImageSource(`${baseUrl}/vector.bin?cache=1`)).resolves
        .toMatchObject({ info: { contentType: 'image/svg+xml' } });
      await expect(resolveImageSource(`${baseUrl}/redirect`)).resolves
        .toMatchObject({ info: { contentType: 'image/svg+xml' } });
      await expect(resolveImageSource(`${baseUrl}/missing`)).rejects.toThrow(/HTTP 404/i);
      await expect(resolveImageSource(`${baseUrl}/malformed`)).rejects.toThrow(TypeError);

      const controller = new AbortController();
      const reason = new Error('abort HTTP SVG source');
      const loading = resolveImageSource(`${baseUrl}/slow`, controller.signal);
      await slowRequested;
      controller.abort(reason);
      await expect(loading).rejects.toBe(reason);
    } finally {
      server.closeAllConnections();
      await new Promise<void>((resolve, reject) => {
        server.close((error) => error ? reject(error) : resolve());
      });
    }
  });

  it('uses Fetch for browser-relative SVG URLs and preserves the input string', async () => {
    const svg = svgBytes();
    const fetchMock = vi.fn(async () => new Response(svg, {
      status: 200,
      headers: { 'content-type': 'text/plain' },
    }));
    vi.stubGlobal('fetch', fetchMock);
    vi.stubGlobal('process', undefined);

    await expect(resolveImageSource('./assets/vector%20one.svg?version=1')).resolves
      .toMatchObject({ info: { contentType: 'image/svg+xml' } });
    expect(fetchMock).toHaveBeenCalledWith(
      './assets/vector%20one.svg?version=1',
      undefined,
    );
  });

  it('aborts before I/O and rejects malformed or unsupported sources', async () => {
    const controller = new AbortController();
    const reason = new Error('abort generic image source');
    controller.abort(reason);
    await expect(resolveImageSource(svgBytes(), controller.signal)).rejects.toBe(reason);

    const invalidSources: unknown[] = [
      new Uint8Array(),
      new Blob(),
      {},
      null,
      [1, 2, 3],
      '',
      'ftp://example.com/vector.svg',
    ];
    for (const source of invalidSources) {
      await expect(resolveImageSource(source as ImageSource)).rejects.toThrow(TypeError);
    }
  });
});
