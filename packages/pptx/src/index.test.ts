import { describe, expect, it } from 'vitest';
import {
  MediaModel,
  PptxDocument,
  inches,
  type ReplaceMediaPosterOptions,
  type ReplaceMediaSourceOptions,
} from './index.js';

describe('@jiayunxie/pptx stable media exports', () => {
  it('runs the complete live media lifecycle through the root package', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const audio = await document.addAudio(0, Uint8Array.of(1, 2, 3), {
      name: 'Root audio',
      altText: 'Root package narration',
      contentType: 'audio/mpeg',
      poster: Uint8Array.of(4),
      posterContentType: 'image/png',
    });
    const video = await document.addVideo(0, 'https://example.com/root-video.mp4');

    expect(audio).toBeInstanceOf(MediaModel);
    expect(document.media(0)[0]).toBe(audio);
    expect(slide.media[0]).toBe(audio);
    expect(slide.shapes[0]).toBe(audio);
    expect(audio.shapeId).toBe(audio.id);
    expect(audio.slidePartUri).toBe(slide.partUri);

    audio.name = 'Root audio edited';
    audio.altText = undefined;
    audio.settings = { play: 'auto', loop: true, volume: 0.5 };
    audio.setTransform({
      x: inches(1),
      y: inches(2),
      width: inches(3),
      height: inches(1),
    });
    const sourceOptions: ReplaceMediaSourceOptions = { contentType: 'audio/wav' };
    expect(await audio.replaceSource(Uint8Array.of(5, 6), sourceOptions)).toBe(audio);
    expect(await audio.replaceSource('https://example.com/root-audio.wav')).toBe(audio);
    expect(await audio.replaceSource(Uint8Array.of(5, 6), sourceOptions)).toBe(audio);
    const posterOptions: ReplaceMediaPosterOptions = { contentType: 'image/gif' };
    expect(await audio.replacePoster(Uint8Array.of(7), posterOptions)).toBe(audio);

    const duplicate = document.duplicateSlide(0);
    const duplicateAudio = duplicate.media[0]!;
    const duplicateVideo = duplicate.media[1]!;
    expect(duplicateAudio).not.toBe(audio);
    expect(duplicateAudio.mediaPartUri).toBe(audio.mediaPartUri);
    await duplicateAudio.replaceSource(Uint8Array.of(8), { contentType: 'audio/ogg' });
    await duplicateAudio.replacePoster(Uint8Array.of(9), { contentType: 'image/jpeg' });
    duplicateVideo.remove();
    document.moveSlide(1, 0);
    expect(document.slides[0]).toBe(duplicate);
    document.moveSlide(0, 1);
    video.remove();
    expect(document.media(0)).toEqual([audio]);
    expect(document.media(1)).toEqual([duplicateAudio]);

    await document.write();
    audio.name = 'Root audio after write';
    const reopened = await PptxDocument.open(await document.write());
    const reopenedAudio = reopened.media(0)[0]!;
    expect(reopenedAudio).toBeInstanceOf(MediaModel);
    expect(reopenedAudio.name).toBe('Root audio after write');
    expect(reopenedAudio.mediaPartUri).not.toBe(reopened.media(1)[0]!.mediaPartUri);
    expect(reopenedAudio.posterPartUri).not.toBe(reopened.media(1)[0]!.posterPartUri);
    await reopenedAudio.replacePoster();
    reopenedAudio.settings = undefined;
    expect(reopenedAudio.settings).toEqual({});
    expect(new TextDecoder().decode(
      reopened.opcPackage.requirePart(reopened.slides[0]!.partUri).bytes,
    )).not.toContain('<p:timing>');
    reopenedAudio.settings = { play: 'click', volume: 1 };
    const second = await PptxDocument.open(await reopened.write());
    expect(second.media(0)[0]!.posterPartUri).toMatch(/\.png$/);
    expect(second.media(0)[0]!.settings).toEqual({
      play: 'click',
      loop: false,
      hideWhenStopped: false,
      volume: 1,
    });
    for (const [slideIndex, expected] of second.slides.entries()) {
      const source = new TextDecoder().decode(second.opcPackage.requirePart(expected.partUri).bytes);
      const ids = [...source.matchAll(/<p:cTn\b[^>]*\bid="([0-9]+)"/g)]
        .map((match) => Number(match[1]));
      const targets = [...source.matchAll(/<p:spTgt\b[^>]*\bspid="([0-9]+)"/g)]
        .map((match) => Number(match[1]));
      expect(source).toContain('cmd="playFrom(0.0)"');
      expect(source).toContain('<p:audio><p:cMediaNode');
      expect(new Set(ids).size).toBe(ids.length);
      expect(new Set(targets)).toEqual(new Set(second.media(slideIndex).map(({ shapeId }) => shapeId)));
    }
    await second.write({ mode: 'permissive' });
    expect(second.diagnostics.filter(({ code }) => code.startsWith('MEDIA_TIMING_'))).toEqual([]);

    if (false) {
      // @ts-expect-error media sources exclude scalar numbers
      await audio.replaceSource(1);
      // @ts-expect-error source replacement excludes placement options
      await audio.replaceSource(Uint8Array.of(1), { x: inches(1) });
      // @ts-expect-error poster replacement excludes transcoders
      await audio.replacePoster(Uint8Array.of(1), { transcode: async () => undefined });
      // @ts-expect-error playback mode excludes hover
      audio.settings = { play: 'hover' };
      // @ts-expect-error media names must be strings
      audio.name = 1;
      // @ts-expect-error shape ids must be numbers
      slide.deleteMedia('2');
    }
  });
});
