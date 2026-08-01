import { createServer } from 'node:http';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, relative } from 'node:path';
import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  resolveSlideBackgroundImage,
  type SetSlideBackgroundImageOptions,
} from './slide-background-source.js';

function pngHeader(width: number, height: number): Uint8Array<ArrayBuffer> {
  const bytes = new Uint8Array(24);
  bytes.set([137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82]);
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

function gifHeader(width: number, height: number): Uint8Array<ArrayBuffer> {
  return Uint8Array.from([
    71, 73, 70, 56, 57, 97,
    width & 0xff, width >>> 8,
    height & 0xff, height >>> 8,
  ]);
}

function jpegHeader(width: number, height: number): Uint8Array<ArrayBuffer> {
  return Uint8Array.from([
    0xff, 0xd8, 0xff, 0xc0, 0x00, 0x08, 0x08,
    height >>> 8, height & 0xff,
    width >>> 8, width & 0xff,
    0x01,
  ]);
}

function dataUri(
  contentType: 'image/png' | 'image/jpeg' | 'image/gif' | 'image/svg+xml',
  bytes: Uint8Array,
): string {
  let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return `data:${contentType};base64,${btoa(binary)}`;
}

afterEach(() => {
  vi.unstubAllGlobals();
});

describe('slide background raster source resolution', () => {
  it('resolves every portable in-memory form into a detached image value', async () => {
    const png = pngHeader(16, 9);
    const gif = gifHeader(8, 6);
    const jpeg = jpegHeader(10, 7);
    const stream = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(png.slice(0, 12));
        controller.enqueue(png.slice(12));
        controller.close();
      },
    });
    const iterable = {
      async *[Symbol.asyncIterator]() {
        yield png.slice(0, 8);
        yield png.slice(8);
      },
    };
    const sources = [
      { source: png, contentType: 'image/png', bytes: png },
      { source: gif.buffer.slice(0), contentType: 'image/gif', bytes: gif },
      { source: new Blob([jpeg], { type: 'text/plain' }), contentType: 'image/jpeg', bytes: jpeg },
      { source: stream, contentType: 'image/png', bytes: png },
      { source: iterable, contentType: 'image/png', bytes: png },
      { source: dataUri('image/png', png), contentType: 'image/png', bytes: png },
    ] as const;

    for (const fixture of sources) {
      const result = await resolveSlideBackgroundImage(fixture.source);
      expect(result).toEqual({
        kind: 'image',
        contentType: fixture.contentType,
        bytes: fixture.bytes,
      });
      expect(Object.isFrozen(result)).toBe(true);
      expect(result.bytes).not.toBe(fixture.bytes);
    }

    const source = pngHeader(4, 3);
    const pending = resolveSlideBackgroundImage(source);
    source.fill(0);
    await expect(pending).resolves.toMatchObject({ bytes: pngHeader(4, 3) });
  });

  it('accepts only closed own-data options and validates MIME before returning', async () => {
    const controller = new AbortController();
    const options = Object.assign(Object.create(null) as Record<string, unknown>, {
      contentType: 'image/png',
      signal: controller.signal,
    }) as SetSlideBackgroundImageOptions;
    await expect(resolveSlideBackgroundImage(pngHeader(1, 1), options)).resolves.toMatchObject({
      kind: 'image',
      contentType: 'image/png',
    });

    let reads = 0;
    const accessor = Object.defineProperty({}, 'contentType', {
      enumerable: true,
      get() {
        reads += 1;
        return 'image/png';
      },
    });
    class Options {}
    const invalid: unknown[] = [
      null,
      [],
      new Options(),
      Object.create({ contentType: 'image/png' }),
      accessor,
      { [Symbol('unsafe')]: true },
      { unknown: true },
      { contentType: 'image/svg+xml' },
      { signal: {} },
    ];
    for (const value of invalid) {
      await expect(resolveSlideBackgroundImage(pngHeader(1, 1), value as never))
        .rejects.toThrow(TypeError);
    }
    expect(reads).toBe(0);

    await expect(resolveSlideBackgroundImage(pngHeader(1, 1), {
      contentType: 'image/gif',
    })).rejects.toThrow(/expected image\/gif.*image\/png/i);
    await expect(resolveSlideBackgroundImage(dataUri('image/gif', pngHeader(1, 1))))
      .rejects.toThrow(/declares image\/gif.*image\/png/i);
  });

  it('propagates raster signature, stream, and abort failures', async () => {
    await expect(resolveSlideBackgroundImage(Uint8Array.of(1, 2, 3)))
      .rejects.toThrow(/signature/i);
    await expect(resolveSlideBackgroundImage(pngHeader(1, 1).slice(0, 8)))
      .rejects.toThrow(/truncated PNG IHDR/i);
    await expect(resolveSlideBackgroundImage(
      dataUri('image/svg+xml', new TextEncoder().encode('<svg/>')),
    )).rejects.toThrow(/cannot declare image\/svg\+xml/i);
    await expect(resolveSlideBackgroundImage({
      async *[Symbol.asyncIterator]() {
        yield 'not bytes';
      },
    } as never)).rejects.toThrow(/streams must yield/i);

    const controller = new AbortController();
    const reason = new Error('abort slide background source');
    controller.abort(reason);
    await expect(resolveSlideBackgroundImage(pngHeader(1, 1), {
      signal: controller.signal,
    })).rejects.toBe(reason);
  });

  it('loads relative and absolute paths plus redirecting query URLs', async () => {
    const png = pngHeader(32, 18);
    const directory = await mkdtemp(join(tmpdir(), 'pptx-background-source-'));
    const path = join(directory, 'background.bin');
    await writeFile(path, png);
    const server = createServer((request, response) => {
      const url = new URL(request.url ?? '/', 'http://127.0.0.1');
      if (url.pathname === '/redirect') {
        response.writeHead(302, { location: '/background?from=redirect' });
        response.end();
        return;
      }
      if (url.pathname === '/background') {
        response.writeHead(200, { 'content-type': 'application/octet-stream' });
        response.end(png);
        return;
      }
      response.writeHead(404);
      response.end();
    });
    await new Promise<void>((resolve, reject) => {
      server.once('error', reject);
      server.listen(0, '127.0.0.1', resolve);
    });
    const address = server.address();
    if (!address || typeof address === 'string') throw new Error('HTTP server has no port');
    const baseUrl = `http://127.0.0.1:${address.port}`;
    try {
      await expect(resolveSlideBackgroundImage(path)).resolves.toMatchObject({ bytes: png });
      await expect(resolveSlideBackgroundImage(relative(process.cwd(), path)))
        .resolves.toMatchObject({ bytes: png });
      await expect(resolveSlideBackgroundImage(`${baseUrl}/background?cache=1`))
        .resolves.toMatchObject({ bytes: png });
      await expect(resolveSlideBackgroundImage(`${baseUrl}/redirect`))
        .resolves.toMatchObject({ bytes: png });
    } finally {
      server.closeAllConnections();
      await new Promise<void>((resolve, reject) => {
        server.close((error) => error ? reject(error) : resolve());
      });
      await rm(directory, { recursive: true, force: true });
    }
  });

  it('uses Fetch for browser-relative sources without rewriting the URL', async () => {
    const png = pngHeader(16, 9);
    const fetchMock = vi.fn(async () => new Response(png, { status: 200 }));
    vi.stubGlobal('fetch', fetchMock);
    vi.stubGlobal('process', undefined);

    await expect(resolveSlideBackgroundImage('./assets/background%20one.png?version=1'))
      .resolves.toMatchObject({ contentType: 'image/png', bytes: png });
    expect(fetchMock).toHaveBeenCalledWith(
      './assets/background%20one.png?version=1',
      undefined,
    );
  });
});
