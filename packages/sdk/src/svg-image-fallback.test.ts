import { afterEach, describe, expect, it, vi } from 'vitest';
import { inspectRasterImage, type ResolvedImageSource } from './raster-image-source.js';
import { resolveSvgFallback } from './svg-image-fallback.js';

const SVG_BYTES = new TextEncoder().encode(
  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 640 360"/>',
);
const VALID_PNG = Uint8Array.from([
  137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82,
  0, 0, 0, 1, 0, 0, 0, 1, 8, 4, 0, 0, 0, 181, 28, 12, 2, 0, 0,
  0, 11, 73, 68, 65, 84, 120, 218, 99, 100, 248, 15, 0, 1, 5, 1, 1,
  39, 24, 227, 102, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
]);

function resolvedSvg(width = 640, height = 360): ResolvedImageSource {
  return {
    bytes: new Uint8Array(SVG_BYTES),
    info: { contentType: 'image/svg+xml', width, height },
  };
}

describe('SVG image fallback selection', () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('preserves and detaches an explicit valid PNG fallback', async () => {
    const fallback = new Uint8Array(VALID_PNG);
    const pending = resolveSvgFallback(resolvedSvg(), fallback);
    fallback.fill(0);
    const resolved = await pending;

    expect(resolved).toEqual(VALID_PNG);
    expect(resolved.buffer).not.toBe(fallback.buffer);
    expect(inspectRasterImage(resolved)).toEqual({
      contentType: 'image/png',
      width: 1,
      height: 1,
    });
  });

  it.each([
    {
      name: 'JPEG',
      bytes: Uint8Array.of(
        0xff, 0xd8, 0xff, 0xc0, 0, 8, 8, 0, 1, 0, 1, 1,
      ),
    },
    { name: 'GIF', bytes: new TextEncoder().encode('GIF89a\x01\x00\x01\x00') },
    { name: 'SVG', bytes: SVG_BYTES },
    { name: 'truncated PNG', bytes: VALID_PNG.slice(0, 8) },
  ])('rejects an explicit $name fallback', async ({ bytes }) => {
    await expect(resolveSvgFallback(resolvedSvg(), bytes))
      .rejects.toThrow(/fallback.*PNG/i);
  });

  it('returns a deterministic detached built-in PNG outside a browser runtime', async () => {
    vi.stubGlobal('Image', undefined);
    vi.stubGlobal('document', undefined);
    const first = await resolveSvgFallback(resolvedSvg(), undefined);
    const second = await resolveSvgFallback(resolvedSvg(), undefined);

    expect(first).toEqual(second);
    expect(first.buffer).not.toBe(second.buffer);
    expect(inspectRasterImage(first)).toEqual({
      contentType: 'image/png',
      width: 1,
      height: 1,
    });
    first.fill(0);
    expect(inspectRasterImage(second).contentType).toBe('image/png');
  });

  it('rasterizes through browser Image and Canvas and cleans the object URL', async () => {
    const drawImage = vi.fn();
    const canvas = {
      width: 0,
      height: 0,
      getContext: vi.fn(() => ({ drawImage })),
      toBlob: vi.fn((callback: BlobCallback) => {
        callback(new Blob([VALID_PNG], { type: 'image/png' }));
      }),
    };
    const createObjectURL = vi.fn(() => 'blob:svg-fallback');
    const revokeObjectURL = vi.fn();
    class BrowserImage {
      src = '';
      decode = vi.fn(async () => undefined);
    }
    vi.stubGlobal('Image', BrowserImage);
    vi.stubGlobal('document', { createElement: vi.fn(() => canvas) });
    vi.stubGlobal('URL', { createObjectURL, revokeObjectURL });

    const fallback = await resolveSvgFallback(resolvedSvg(640.2, 359.1), undefined);

    expect(fallback).toEqual(VALID_PNG);
    expect(createObjectURL).toHaveBeenCalledOnce();
    expect(canvas.width).toBe(641);
    expect(canvas.height).toBe(360);
    expect(drawImage).toHaveBeenCalledWith(expect.any(BrowserImage), 0, 0, 641, 360);
    expect(canvas.toBlob).toHaveBeenCalledOnce();
    expect(revokeObjectURL).toHaveBeenCalledWith('blob:svg-fallback');
  });

  it.each([
    { name: 'decode failure', mode: 'decode' },
    { name: 'missing context', mode: 'context' },
    { name: 'tainted draw', mode: 'draw' },
    { name: 'null export', mode: 'export' },
  ])('uses the built-in PNG after $name and still revokes the URL', async ({ mode }) => {
    const revokeObjectURL = vi.fn();
    class BrowserImage {
      src = '';
      async decode(): Promise<void> {
        if (mode === 'decode') throw new Error('decode failed');
      }
    }
    const context = mode === 'context'
      ? null
      : {
          drawImage() {
            if (mode === 'draw') throw new Error('tainted canvas');
          },
        };
    const canvas = {
      width: 0,
      height: 0,
      getContext: () => context,
      toBlob: (callback: BlobCallback) => callback(
        mode === 'export' ? null : new Blob([VALID_PNG], { type: 'image/png' }),
      ),
    };
    vi.stubGlobal('Image', BrowserImage);
    vi.stubGlobal('document', { createElement: () => canvas });
    vi.stubGlobal('URL', {
      createObjectURL: () => 'blob:failed-svg',
      revokeObjectURL,
    });

    const fallback = await resolveSvgFallback(resolvedSvg(), undefined);

    expect(fallback).toEqual(VALID_PNG);
    expect(revokeObjectURL).toHaveBeenCalledWith('blob:failed-svg');
  });

  it.each([
    { width: 8193, height: 1 },
    { width: 4097, height: 4097 },
    { width: Number.MAX_VALUE, height: 1 },
  ])('uses the built-in PNG when $width x $height exceeds the canvas budget', async ({
    width,
    height,
  }) => {
    const createObjectURL = vi.fn();
    vi.stubGlobal('Image', class {});
    vi.stubGlobal('document', { createElement: vi.fn() });
    vi.stubGlobal('URL', { createObjectURL, revokeObjectURL: vi.fn() });

    const fallback = await resolveSvgFallback(resolvedSvg(width, height), undefined);

    expect(fallback).toEqual(VALID_PNG);
    expect(createObjectURL).not.toHaveBeenCalled();
  });

  it('propagates abort during image decode and revokes the object URL', async () => {
    const controller = new AbortController();
    const reason = new Error('abort SVG rasterization');
    const revokeObjectURL = vi.fn();
    class PendingImage {
      src = '';
      decode(): Promise<void> {
        return new Promise(() => undefined);
      }
    }
    vi.stubGlobal('Image', PendingImage);
    vi.stubGlobal('document', { createElement: vi.fn() });
    vi.stubGlobal('URL', {
      createObjectURL: () => 'blob:pending-svg',
      revokeObjectURL,
    });

    const pending = resolveSvgFallback(resolvedSvg(), undefined, controller.signal);
    controller.abort(reason);

    await expect(pending).rejects.toBe(reason);
    expect(revokeObjectURL).toHaveBeenCalledWith('blob:pending-svg');
  });
});
