import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it, vi } from 'vitest';
import { normalizeMediaCreateRequest } from './media-create.internal.js';
import { resolveMediaCreationInputs } from './media-source.internal.js';

describe('media creation source resolution', () => {
  it('resolves canonical media and poster data URIs', async () => {
    const request = normalizeMediaCreateRequest(
      'audio',
      'data:audio/mpeg;base64,AQIDBA==',
      {
        contentType: 'audio/mpeg',
        fileName: 'voice.mp3',
        poster: 'data:image/png;base64,iVBORw0KGgo=',
        posterContentType: 'image/png',
      },
    );
    const resolved = await resolveMediaCreationInputs(request);

    expect(resolved).toEqual({
      media: {
        type: 'embedded',
        bytes: new Uint8Array([1, 2, 3, 4]),
        contentType: 'audio/mpeg',
        extension: '.mp3',
      },
      poster: {
        type: 'embedded',
        bytes: new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]),
        contentType: 'image/png',
        extension: '.png',
      },
    });
    expect(Object.isFrozen(resolved)).toBe(true);
    expect(Object.isFrozen(resolved.media)).toBe(true);
    expect(Object.isFrozen(resolved.poster)).toBe(true);
    expect(Object.isFrozen(resolved.poster.bytes)).toBe(false);
  });

  it('infers and preserves safe path and File extensions', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'pptx-media-source-'));
    const oggPath = join(directory, 'voice.ogg');
    const m4vPath = join(directory, 'clip.m4v');
    const jpegPath = join(directory, 'cover.jpeg');
    await writeFile(oggPath, Uint8Array.of(1, 2, 3));
    await writeFile(m4vPath, Uint8Array.of(4, 5));
    await writeFile(jpegPath, Uint8Array.of(6, 7));
    try {
      const fromPath = await resolveMediaCreationInputs(
        normalizeMediaCreateRequest('audio', oggPath, {}),
      );
      expect(fromPath.media).toMatchObject({
        type: 'embedded',
        bytes: new Uint8Array([1, 2, 3]),
        contentType: 'audio/ogg',
        extension: '.ogg',
      });

      const fromNamedPaths = await resolveMediaCreationInputs(
        normalizeMediaCreateRequest('video', m4vPath, { poster: jpegPath }),
      );
      expect(fromNamedPaths.media).toMatchObject({
        contentType: 'video/mp4',
        extension: '.m4v',
      });
      expect(fromNamedPaths.poster).toMatchObject({
        contentType: 'image/jpeg',
        extension: '.jpeg',
      });

      if (typeof File !== 'undefined') {
        const file = new File([Uint8Array.of(4, 5)], 'clip.m4v', { type: 'text/plain' });
        const poster = new File([Uint8Array.of(6, 7)], 'cover.jpeg', { type: 'text/plain' });
        const fromFiles = await resolveMediaCreationInputs(
          normalizeMediaCreateRequest('video', file, { poster }),
        );
        expect(fromFiles.media).toMatchObject({
          contentType: 'video/mp4',
          extension: '.m4v',
        });
        expect(fromFiles.poster).toMatchObject({
          contentType: 'image/jpeg',
          extension: '.jpeg',
        });
      }
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  });

  it('resolves bytes, Blob, Web Stream, and async iterable sources', async () => {
    const iterable = {
      async *[Symbol.asyncIterator]() {
        yield Uint8Array.of(1, 2);
        yield new DataView(Uint8Array.of(3, 4).buffer);
      },
    };
    const stream = new ReadableStream<number>({
      start(controller) {
        controller.enqueue(5);
        controller.enqueue(6);
        controller.close();
      },
    });
    const cases = [
      normalizeMediaCreateRequest('video', Uint8Array.of(1), {
        contentType: 'video/webm',
      }),
      normalizeMediaCreateRequest('video', Uint8Array.of(2).buffer, {
        contentType: 'video/quicktime',
      }),
      normalizeMediaCreateRequest('video', new Blob([Uint8Array.of(3)]), {
        contentType: 'video/mp4',
      }),
      normalizeMediaCreateRequest('audio', iterable, {
        contentType: 'audio/wav',
      }),
      normalizeMediaCreateRequest('audio', stream, {
        contentType: 'audio/mp4',
      }),
    ];
    const resolved = await Promise.all(cases.map(resolveMediaCreationInputs));

    expect(resolved.map(({ media }) => media.type === 'embedded'
      ? [media.contentType, media.extension, [...media.bytes]]
      : [])).toEqual([
      ['video/webm', '.webm', [1]],
      ['video/quicktime', '.mov', [2]],
      ['video/mp4', '.mp4', [3]],
      ['audio/wav', '.wav', [1, 2, 3, 4]],
      ['audio/mp4', '.m4a', [5, 6]],
    ]);

    const defaulted = await resolveMediaCreationInputs(
      normalizeMediaCreateRequest('audio', Uint8Array.of(7), {
        fileName: 'voice.unknown',
      }),
    );
    expect(defaulted.media).toMatchObject({
      contentType: 'audio/mpeg',
      extension: '.mp3',
    });

    const sharedBuffer = Uint8Array.of(8, 9).buffer;
    const mutatingIterable = {
      async *[Symbol.asyncIterator]() {
        yield sharedBuffer;
        new Uint8Array(sharedBuffer)[0] = 0;
        yield 10;
      },
    };
    const detachedChunks = await resolveMediaCreationInputs(
      normalizeMediaCreateRequest('audio', mutatingIterable, {}),
    );
    expect(detachedChunks.media).toMatchObject({ bytes: Uint8Array.of(8, 9, 10) });
  });

  it('keeps HTTP and HTTPS media external and supplies a detached default poster', async () => {
    const first = await resolveMediaCreationInputs(
      normalizeMediaCreateRequest('video', 'https://example.com/media/video.mp4?x=1', {}),
    );
    const second = await resolveMediaCreationInputs(
      normalizeMediaCreateRequest('audio', 'HTTP://example.com/audio.mp3', {}),
    );

    expect(first.media).toEqual({
      type: 'external',
      url: 'https://example.com/media/video.mp4?x=1',
    });
    expect(second.media).toEqual({
      type: 'external',
      url: 'HTTP://example.com/audio.mp3',
    });
    expect(first.poster).toMatchObject({ contentType: 'image/png', extension: '.png' });
    expect(first.poster.bytes.length).toBeGreaterThan(8);
    expect(first.poster.bytes).not.toBe(second.poster.bytes);
    first.poster.bytes[0] = 0;
    expect(second.poster.bytes[0]).toBe(137);
  });

  it('rejects non-canonical or malformed data URIs', async () => {
    const invalid = [
      'data:audio/mpeg;base64,',
      'data:audio/mpeg,AQ==',
      'data:audio/mpeg;base64,AQ',
      'data:audio/mpeg;base64,AB==',
      'data:audio/mpeg;base64,AQ===',
      'data:audio/mpeg;base64,AQ==\n',
      'data:audio/mpeg;base64,AQ%3D%3D',
      'data:audio/mpeg;base64,AQ-_',
      'data:audio/mpeg;base64,AQ==,AA==',
      'data:Audio/Mpeg;base64,AQ==',
      'data:audio/flac;base64,AQ==',
      'data:image/png;base64,AQ==',
    ];
    for (const source of invalid) {
      await expect(resolveMediaCreationInputs(
        normalizeMediaCreateRequest('audio', source, {}),
      )).rejects.toThrow();
    }
  });

  it('rejects MIME, extension, scheme, poster, and empty-payload conflicts before stream I/O', async () => {
    const consumed = vi.fn();
    const source = {
      async *[Symbol.asyncIterator]() {
        consumed();
        yield Uint8Array.of(1);
      },
    };
    const invalid = [
      normalizeMediaCreateRequest('audio', source, { contentType: 'Audio/Mpeg' }),
      normalizeMediaCreateRequest('audio', source, { contentType: 'video/mp4' }),
      normalizeMediaCreateRequest('audio', source, { contentType: 'audio/flac' }),
      normalizeMediaCreateRequest('audio', source, {
        contentType: 'audio/mpeg',
        fileName: 'voice.wav',
      }),
      normalizeMediaCreateRequest('audio', 'data:audio/mpeg;base64,AQ==', {
        contentType: 'audio/wav',
      }),
      normalizeMediaCreateRequest('audio', source, {
        poster: 'https://example.com/cover.png',
      }),
      normalizeMediaCreateRequest('audio', source, {
        poster: new Blob([]),
      }),
      normalizeMediaCreateRequest('audio', source, {
        posterContentType: 'image/jpeg',
      }),
    ];
    for (const request of invalid) {
      await expect(resolveMediaCreationInputs(request)).rejects.toThrow();
    }
    expect(consumed).not.toHaveBeenCalled();

    await expect(resolveMediaCreationInputs(
      normalizeMediaCreateRequest('audio', 'ftp://example.com/voice.mp3', {}),
    )).rejects.toThrow(/scheme/i);
    await expect(resolveMediaCreationInputs(
      normalizeMediaCreateRequest('audio', new Blob([]), {}),
    )).rejects.toThrow(/empty/i);
    await expect(resolveMediaCreationInputs(
      normalizeMediaCreateRequest('audio', {
        async *[Symbol.asyncIterator]() {
          yield 'not bytes';
        },
      }, {}),
    )).rejects.toThrow(/byte/i);
  });

  it('applies a strict detached embedded transcoder result', async () => {
    const output = Uint8Array.of(8, 9);
    const result = { bytes: output, contentType: 'audio/ogg', extension: 'ogg' };
    const transcode = vi.fn(async (bytes: Uint8Array, contentType: string, kind: string) => {
      expect(bytes).toEqual(Uint8Array.of(1, 2, 3));
      expect(contentType).toBe('audio/mpeg');
      expect(kind).toBe('audio');
      bytes[0] = 0;
      return result;
    });
    const resolved = await resolveMediaCreationInputs(
      normalizeMediaCreateRequest('audio', Uint8Array.of(1, 2, 3), {
        contentType: 'audio/mpeg',
        transcode,
      }),
    );

    output[0] = 0;
    result.contentType = 'audio/wav';
    expect(transcode).toHaveBeenCalledOnce();
    expect(resolved.media).toEqual({
      type: 'embedded',
      bytes: Uint8Array.of(8, 9),
      contentType: 'audio/ogg',
      extension: '.ogg',
    });
  });

  it('rejects invalid or external transcode results without invoking unsafe accessors', async () => {
    const externalTranscode = vi.fn(async () => ({
      bytes: Uint8Array.of(1),
      contentType: 'video/mp4',
    }));
    await expect(resolveMediaCreationInputs(
      normalizeMediaCreateRequest('video', 'https://example.com/video.mp4', {
        transcode: externalTranscode,
      }),
    )).rejects.toThrow(/external/i);
    expect(externalTranscode).not.toHaveBeenCalled();

    const getter = vi.fn(() => Uint8Array.of(1));
    const accessor = { contentType: 'audio/mpeg' };
    Object.defineProperty(accessor, 'bytes', { enumerable: true, get: getter });
    const invalid: unknown[] = [
      null,
      [],
      new (class Result { bytes = Uint8Array.of(1); contentType = 'audio/mpeg'; })(),
      accessor,
      { bytes: new Uint8Array(), contentType: 'audio/mpeg' },
      { bytes: Uint8Array.of(1), contentType: 'video/mp4' },
      { bytes: Uint8Array.of(1), contentType: 'audio/mpeg', extension: '.wav' },
      { bytes: Uint8Array.of(1), contentType: 'audio/mpeg', extension: '.MP3' },
      { bytes: Uint8Array.of(1), contentType: 'audio/mpeg', unknown: true },
    ];
    for (const result of invalid) {
      await expect(resolveMediaCreationInputs(
        normalizeMediaCreateRequest('audio', Uint8Array.of(1), {
          transcode: async () => result as never,
        }),
      )).rejects.toThrow();
    }
    expect(getter).not.toHaveBeenCalled();
  });
});
