import { describe, expect, it } from 'vitest';
import { ImageModel, inches } from '@pptx/model';
import { PptxDocument } from './index.js';
import {
  commitPreparedImage,
  prepareImageSource,
} from './prepared-image-source.js';

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

function svg(width = 640, height = 360): Uint8Array<ArrayBuffer> {
  return new TextEncoder().encode(
    `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${width} ${height}"/>`,
  );
}

describe('prepared image sources', () => {
  it('resolves once, detaches bytes, and commits page-local relationships to one payload', async () => {
    const bytes = pngHeader(16, 9);
    let reads = 0;
    const source = {
      async *[Symbol.asyncIterator]() {
        reads += 1;
        yield bytes;
      },
    };
    const prepared = await prepareImageSource(source, {
      x: inches(0.5),
      y: inches(0.75),
      width: inches(2),
      height: inches(1),
    });
    bytes.fill(0);

    const deck = PptxDocument.create();
    const first = deck.addSlide();
    const second = deck.addSlide();
    const firstImage = commitPreparedImage(first, prepared);
    const secondImage = commitPreparedImage(second, prepared);

    expect(reads).toBe(1);
    expect(firstImage).toBeInstanceOf(ImageModel);
    expect(firstImage).not.toBe(secondImage);
    expect(firstImage.transform).toMatchObject({
      x: inches(0.5),
      y: inches(0.75),
      width: inches(2),
      height: inches(1),
    });
    expect(firstImage.sourcePartUri).toBe(secondImage.sourcePartUri);
    expect(first.relationships.filter(({ resolvedTarget }) =>
      resolvedTarget === firstImage.sourcePartUri)).toHaveLength(1);
    expect(second.relationships.filter(({ resolvedTarget }) =>
      resolvedTarget === secondImage.sourcePartUri)).toHaveLength(1);
    expect(deck.opcPackage.requirePart(firstImage.sourcePartUri!).bytes)
      .toEqual(pngHeader(16, 9));
  });

  it('prepares one SVG fallback pair and reuses both exact payloads', async () => {
    const prepared = await prepareImageSource(svg(), {
      fallback: pngHeader(1, 1),
      width: inches(3),
      height: inches(2),
    });
    const deck = PptxDocument.create();
    const firstImage = commitPreparedImage(deck.addSlide(), prepared);
    const secondImage = commitPreparedImage(deck.addSlide(), prepared);

    expect(firstImage.isSvg).toBe(true);
    expect(secondImage.isSvg).toBe(true);
    expect(firstImage.fallbackPartUri).toBe(secondImage.fallbackPartUri);
    expect(firstImage.svgPartUri).toBe(secondImage.svgPartUri);
    expect(firstImage.fallbackPartUri).not.toBe(firstImage.svgPartUri);
  });

  it('calculates sizing independently from each page placeholder', async () => {
    const deck = PptxDocument.create();
    const layout = deck.layouts[0]!;
    const prompt = layout.addPlaceholder('Picture', {
      name: 'hero',
      type: 'pic',
      index: 301,
      x: inches(1),
      y: inches(1.5),
      width: inches(4),
      height: inches(2),
    });
    const prepared = await prepareImageSource(pngHeader(16, 9), {
      placeholder: 'hero',
      sizing: { type: 'cover', width: inches(9), height: inches(9) },
    });
    const firstImage = commitPreparedImage(deck.addSlide({ masterName: 'DEFAULT' }), prepared);
    const secondImage = commitPreparedImage(deck.addSlide({ masterName: 'DEFAULT' }), prepared);

    expect(firstImage.transform).toEqual(prompt.transform);
    expect(secondImage.transform).toEqual(prompt.transform);
    expect(firstImage.sourceRectangle).toEqual({
      left: 0,
      top: 5.556,
      right: 0,
      bottom: 5.556,
    });
  });

  it('rejects invalid fallback and aborted loads before a document is mutated', async () => {
    const deck = PptxDocument.create();
    const before = deck.opcPackage.parts;
    await expect(prepareImageSource(pngHeader(1, 1), {
      fallback: pngHeader(1, 1),
    })).rejects.toThrow(/fallback is only valid for SVG/i);

    const controller = new AbortController();
    const reason = new Error('stop image preparation');
    controller.abort(reason);
    await expect(prepareImageSource(pngHeader(1, 1), {
      signal: controller.signal,
    })).rejects.toBe(reason);
    expect(deck.opcPackage.parts).toEqual(before);
    expect(deck.slides).toEqual([]);
  });
});
