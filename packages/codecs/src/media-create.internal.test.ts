import { describe, expect, it, vi } from 'vitest';
import type { AddMediaOptions, MediaByteStream } from './media.js';
import {
  finalizeMediaCreationDefinition,
  normalizeMediaCreateRequest,
  renderMediaPictureXml,
} from './media-create.internal.js';
import { resolveMediaCreationInputs } from './media-source.internal.js';

describe('media creation request normalization', () => {
  it('publishes strict name, alt-text, and placeholder option types', () => {
    const valid: AddMediaOptions = {
      name: 'Narration',
      altText: '',
      placeholder: { type: 'media', index: 7 },
    };
    if (false) {
      // @ts-expect-error media name must be a string
      const invalidName: AddMediaOptions = { name: 1 };
      // @ts-expect-error media alt text must be a string
      const invalidAltText: AddMediaOptions = { altText: false };
      void [invalidName, invalidAltText];
    }
    expect(valid).toEqual({
      name: 'Narration',
      altText: '',
      placeholder: { type: 'media', index: 7 },
    });
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
    const placeholder = { type: 'media' as const, index: 12 };
    const options: AddMediaOptions = {
      name: '',
      altText: '',
      placeholder,
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
    placeholder.index = 99;
    expect(request).toMatchObject({
      kind: 'audio',
      source: { type: 'string', value: 'voice.mp3' },
      poster: { type: 'blob' },
      name: '',
      altText: '',
      placeholder: { type: 'media', index: 12 },
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
    expect(Object.isFrozen(request.placeholder)).toBe(true);

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
    const placeholderAccessor = {};
    const placeholderGetter = vi.fn(() => 'media');
    Object.defineProperty(placeholderAccessor, 'type', {
      enumerable: true,
      get: placeholderGetter,
    });
    Object.defineProperty(placeholderAccessor, 'index', {
      enumerable: true,
      value: 1,
    });
    const invalid: unknown[] = [
      { name: 1 },
      { name: 'bad\u0000name' },
      { altText: 'bad\uD800text' },
      { placeholder: '' },
      { placeholder: 'bad\u0000name' },
      { placeholder: null },
      { placeholder: [] },
      { placeholder: { type: 'media' } },
      { placeholder: { type: 'media', index: 1, extra: true } },
      { placeholder: { type: 'audio', index: 1 } },
      { placeholder: { type: 'media', index: -1 } },
      { placeholder: { type: 'media', index: 4_294_967_295 } },
      { placeholder: { type: 'media', index: 1.5 } },
      { placeholder: placeholderAccessor },
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
    expect(placeholderGetter).not.toHaveBeenCalled();
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

describe('media creation definition and XML', () => {
  it('finalizes a detached frozen embedded definition', async () => {
    const mediaBytes = Uint8Array.of(1, 2, 3);
    const posterBytes = Uint8Array.of(4, 5, 6);
    const request = normalizeMediaCreateRequest('audio', mediaBytes, {
      name: 'Narration',
      altText: 'Spoken overview',
      contentType: 'audio/mpeg',
      poster: posterBytes,
      posterContentType: 'image/png',
      x: -1,
      y: 0,
      width: 1,
      height: 2,
      play: 'auto',
      loop: true,
      hideWhenStopped: true,
      volume: 0.25,
    });
    const resolved = await resolveMediaCreationInputs(request);
    const definition = finalizeMediaCreationDefinition(request, resolved, 'Media 0');

    mediaBytes[0] = 9;
    posterBytes[0] = 9;
    if (resolved.media.type === 'embedded') resolved.media.bytes[0] = 8;
    resolved.poster.bytes[0] = 8;
    expect(definition).toEqual({
      kind: 'audio',
      name: 'Narration',
      altText: 'Spoken overview',
      x: -1,
      y: 0,
      width: 1,
      height: 2,
      rotation: 0,
      flipHorizontal: false,
      flipVertical: false,
      play: 'auto',
      loop: true,
      hideWhenStopped: true,
      volume: 0.25,
      media: {
        type: 'embedded',
        bytes: Uint8Array.of(1, 2, 3),
        contentType: 'audio/mpeg',
        extension: '.mp3',
      },
      poster: {
        type: 'embedded',
        bytes: Uint8Array.of(4, 5, 6),
        contentType: 'image/png',
        extension: '.png',
      },
    });
    expect(Object.isFrozen(definition)).toBe(true);
    expect(Object.isFrozen(definition.media)).toBe(true);
    expect(Object.isFrozen(definition.poster)).toBe(true);
    expect(Object.isFrozen(definition.poster.bytes)).toBe(false);
  });

  it('uses a validated default name and preserves omitted alt text', async () => {
    const request = normalizeMediaCreateRequest(
      'video',
      'https://example.com/video.mp4',
      {},
    );
    const resolved = await resolveMediaCreationInputs(request);
    const definition = finalizeMediaCreationDefinition(request, resolved, 'Media 1');

    expect(definition.name).toBe('Media 1');
    expect(definition.altText).toBeUndefined();
    expect(definition.media).toEqual({
      type: 'external',
      url: 'https://example.com/video.mp4',
    });
    expect(() => finalizeMediaCreationDefinition(request, resolved, 'bad\u0000name'))
      .toThrow(/XML/i);
  });

  it('renders exact canonical embedded audio picture XML', async () => {
    const request = normalizeMediaCreateRequest('audio', Uint8Array.of(1), {
      name: 'Audio & narration',
      altText: 'Spoken "overview"',
      contentType: 'audio/mpeg',
      poster: Uint8Array.of(2),
      posterContentType: 'image/png',
      x: -1,
      y: 0,
      width: 1,
      height: 2,
    });
    const definition = finalizeMediaCreationDefinition(
      request,
      await resolveMediaCreationInputs(request),
      'Media 0',
    );
    const xml = renderMediaPictureXml(2, definition, {
      kind: 'rId2',
      media: 'rId3',
      poster: 'rId4',
    });

    expect(xml).toBe(
      '<p:pic><p:nvPicPr><p:cNvPr id="2" name="Audio &amp; narration" descr="Spoken &quot;overview&quot;">'
      + '<a:hlinkClick r:id="" action="ppaction://media"/></p:cNvPr>'
      + '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr>'
      + '<a:audioFile r:link="rId2"/><p:extLst>'
      + '<p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}">'
      + '<p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="rId3"/>'
      + '</p:ext><p:ext uri="{C13D3E4A-5148-4B6D-A7E7-505054582D4F}">'
      + '<px:playback xmlns:px="urn:pptx-ooxml:media" play="click" loop="0" '
      + 'hideWhenStopped="0" volume="100000"/></p:ext></p:extLst></p:nvPr></p:nvPicPr>'
      + '<p:blipFill><a:blip r:embed="rId4"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
      + '<p:spPr><a:xfrm><a:off x="-1" y="0"/><a:ext cx="1" cy="2"/></a:xfrm>'
      + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>',
    );
  });

  it('renders video, external media, escaped ids, and alt-text presence exactly', async () => {
    const embeddedRequest = normalizeMediaCreateRequest('video', Uint8Array.of(1), {
      name: '',
      altText: '',
      contentType: 'video/mp4',
    });
    const embedded = finalizeMediaCreationDefinition(
      embeddedRequest,
      await resolveMediaCreationInputs(embeddedRequest),
      'Media 0',
    );
    const embeddedXml = renderMediaPictureXml(3, embedded, {
      kind: 'kind&1',
      media: 'media"1',
      poster: 'poster<1',
    });
    expect(embeddedXml).toContain('<p:cNvPr id="3" name="" descr="">');
    expect(embeddedXml).toContain('<a:videoFile r:link="kind&amp;1"/>');
    expect(embeddedXml).toContain('r:embed="media&quot;1"');
    expect(embeddedXml).toContain('<a:blip r:embed="poster&lt;1"/>');

    const externalRequest = normalizeMediaCreateRequest(
      'video',
      'https://example.com/video.mp4',
      { play: 'auto', loop: true, hideWhenStopped: true, volume: 0 },
    );
    const external = finalizeMediaCreationDefinition(
      externalRequest,
      await resolveMediaCreationInputs(externalRequest),
      'Media 1',
    );
    const externalXml = renderMediaPictureXml(4, external, {
      kind: 'rId5',
      poster: 'rId6',
    });
    expect(externalXml).toContain('<p:cNvPr id="4" name="Media 1">');
    expect(externalXml).not.toContain(' descr=');
    expect(externalXml).toContain('<a:videoFile r:link="rId5"/>');
    expect(externalXml).not.toContain('p14:media');
    expect(externalXml).toContain(
      'play="auto" loop="1" hideWhenStopped="1" volume="0"',
    );
    expect(externalXml).toContain('<a:picLocks noChangeAspect="1"/>');
  });

  it('rejects inconsistent relationship ids and shape ids', async () => {
    const embeddedRequest = normalizeMediaCreateRequest(
      'audio',
      Uint8Array.of(1),
      { contentType: 'audio/mpeg' },
    );
    const embedded = finalizeMediaCreationDefinition(
      embeddedRequest,
      await resolveMediaCreationInputs(embeddedRequest),
      'Media 0',
    );
    expect(() => renderMediaPictureXml(0, embedded, {
      kind: 'rId1',
      media: 'rId2',
      poster: 'rId3',
    })).toThrow(/shape id/i);
    expect(() => renderMediaPictureXml(2, embedded, {
      kind: 'rId1',
      poster: 'rId3',
    })).toThrow(/media relationship/i);
    expect(() => renderMediaPictureXml(2, embedded, {
      kind: 'bad\u0000id',
      media: 'rId2',
      poster: 'rId3',
    })).toThrow(/XML/i);

    const externalRequest = normalizeMediaCreateRequest(
      'video',
      'https://example.com/video.mp4',
      {},
    );
    const external = finalizeMediaCreationDefinition(
      externalRequest,
      await resolveMediaCreationInputs(externalRequest),
      'Media 1',
    );
    expect(() => renderMediaPictureXml(2, external, {
      kind: 'rId1',
      media: 'rId2',
      poster: 'rId3',
    })).toThrow(/external/i);
  });
});
