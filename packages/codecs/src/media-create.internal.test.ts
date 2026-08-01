import { describe, expect, it, vi } from 'vitest';
import type { AddMediaOptions, MediaByteStream } from './media.js';
import { normalizeMediaCreateRequest } from './media-create.internal.js';

describe('media creation request normalization', () => {
  it('publishes strict name and alt-text option types', () => {
    const valid: AddMediaOptions = { name: 'Narration', altText: '' };
    if (false) {
      // @ts-expect-error media name must be a string
      const invalidName: AddMediaOptions = { name: 1 };
      // @ts-expect-error media alt text must be a string
      const invalidAltText: AddMediaOptions = { altText: false };
      void [invalidName, invalidAltText];
    }
    expect(valid).toEqual({ name: 'Narration', altText: '' });
  });

  it('normalizes defaults and detaches direct byte sources', () => {
    const media = new Uint8Array([1, 2, 3]);
    const poster = new Uint8Array([4, 5, 6]);
    const request = normalizeMediaCreateRequest('audio', media, { poster });

    media[0] = 9;
    poster[0] = 9;
    expect(request).toMatchObject({
      kind: 'audio',
      x: 914_400,
      y: 914_400,
      width: 914_400,
      height: 914_400,
      play: 'click',
      loop: false,
      hideWhenStopped: false,
      volume: 1,
    });
    expect(request.name).toBeUndefined();
    expect(request.altText).toBeUndefined();
    expect(request.source).toMatchObject({ type: 'bytes', bytes: new Uint8Array([1, 2, 3]) });
    expect(request.poster).toMatchObject({ type: 'bytes', bytes: new Uint8Array([4, 5, 6]) });
    expect(Object.isFrozen(request)).toBe(true);
    expect(Object.isFrozen(request.source)).toBe(true);
    expect(Object.isFrozen(request.poster)).toBe(true);
    expect(Object.isFrozen((request.source as { bytes: Uint8Array }).bytes)).toBe(false);

    const buffer = Uint8Array.from([7, 8, 9]).buffer;
    const fromBuffer = normalizeMediaCreateRequest('video', buffer, {});
    new Uint8Array(buffer)[0] = 0;
    expect(fromBuffer).toMatchObject({
      kind: 'video',
      width: 4_572_000,
      height: 2_571_750,
      source: { type: 'bytes', bytes: new Uint8Array([7, 8, 9]) },
    });
  });

  it('preserves every valid explicit option without retaining the options object', () => {
    const transcode: NonNullable<AddMediaOptions['transcode']> = async (bytes, contentType) => ({
      bytes,
      contentType,
    });
    const options: AddMediaOptions = {
      name: '',
      altText: '',
      contentType: 'audio/mpeg',
      fileName: 'voice.mp3',
      poster: new Blob([Uint8Array.of(1)], { type: 'image/png' }),
      posterContentType: 'image/png',
      x: -1,
      y: 0,
      width: 1,
      height: 2,
      play: 'auto',
      loop: true,
      hideWhenStopped: true,
      volume: 0,
      transcode,
    };
    const request = normalizeMediaCreateRequest('audio', 'voice.mp3', options);

    (options as { name?: string }).name = 'changed';
    expect(request).toMatchObject({
      kind: 'audio',
      source: { type: 'string', value: 'voice.mp3' },
      poster: { type: 'blob' },
      name: '',
      altText: '',
      contentType: 'audio/mpeg',
      fileName: 'voice.mp3',
      posterContentType: 'image/png',
      x: -1,
      y: 0,
      width: 1,
      height: 2,
      play: 'auto',
      loop: true,
      hideWhenStopped: true,
      volume: 0,
      transcode,
    });

    const nullPrototype = Object.assign(Object.create(null) as Record<string, unknown>, {
      x: -0,
      y: -0,
      width: 1,
      height: 1,
    });
    const normalizedZero = normalizeMediaCreateRequest('video', Uint8Array.of(1), nullPrototype);
    expect(Object.is(normalizedZero.x, -0)).toBe(false);
    expect(Object.is(normalizedZero.y, -0)).toBe(false);
  });

  it('retains stream handles without consuming them', () => {
    const consumed = vi.fn();
    const iterable: MediaByteStream = {
      async *[Symbol.asyncIterator]() {
        consumed();
        yield Uint8Array.of(1);
      },
    };
    const stream = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(Uint8Array.of(2));
        controller.close();
      },
    });

    const iterableRequest = normalizeMediaCreateRequest('audio', iterable, {});
    const streamRequest = normalizeMediaCreateRequest('video', stream, {});
    expect(iterableRequest.source).toEqual({ type: 'stream', value: iterable });
    expect(streamRequest.source).toEqual({ type: 'stream', value: stream });
    expect(consumed).not.toHaveBeenCalled();
    expect(stream.locked).toBe(false);
  });

  it('rejects unsafe option containers without invoking accessors or consuming sources', () => {
    const getter = vi.fn(() => 'unsafe');
    const accessor = {};
    Object.defineProperty(accessor, 'name', { enumerable: true, get: getter });
    const consumed = vi.fn();
    const source: MediaByteStream = {
      async *[Symbol.asyncIterator]() {
        consumed();
        yield Uint8Array.of(1);
      },
    };
    class Options {
      x = 0;
    }

    const invalid = [
      null,
      undefined,
      true,
      1,
      'options',
      [],
      new Options(),
      Object.create({ x: 0 }),
      accessor,
      { unknown: true },
      { [Symbol('option')]: true },
    ];
    for (const options of invalid) {
      expect(() => normalizeMediaCreateRequest('audio', source, options)).toThrow(TypeError);
    }
    expect(getter).not.toHaveBeenCalled();
    expect(consumed).not.toHaveBeenCalled();
  });

  it('rejects invalid scalar options synchronously', () => {
    const invalid: unknown[] = [
      { name: 1 },
      { name: 'bad\u0000name' },
      { altText: 'bad\uD800text' },
      { contentType: '' },
      { contentType: 1 },
      { posterContentType: '' },
      { fileName: '' },
      { poster: null },
      { transcode: true },
      { x: 1.5 },
      { x: Number.MAX_SAFE_INTEGER + 1 },
      { y: Number.NaN },
      { width: 0 },
      { height: -1 },
      { play: 'hover' },
      { loop: 1 },
      { hideWhenStopped: 'false' },
      { volume: -0.01 },
      { volume: 1.01 },
      { volume: Number.POSITIVE_INFINITY },
    ];
    for (const options of invalid) {
      expect(() => normalizeMediaCreateRequest('audio', Uint8Array.of(1), options)).toThrow();
    }
  });

  it('rejects unsupported or empty source values', () => {
    for (const source of [
      '',
      new Uint8Array(),
      new ArrayBuffer(0),
      null,
      undefined,
      1,
      {},
      [],
      new DataView(new ArrayBuffer(1)),
    ]) {
      expect(() => normalizeMediaCreateRequest('audio', source, {})).toThrow();
    }
    expect(() => normalizeMediaCreateRequest('music' as never, Uint8Array.of(1), {})).toThrow(TypeError);
  });
});
