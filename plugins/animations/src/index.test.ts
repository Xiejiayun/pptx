import { describe, expect, it } from 'vitest';
import { CodecRegistry, MediaCodec } from '@pptx/codecs';
import { LosslessXmlDocument } from '@pptx/lossless-xml';
import { OpcPackage } from '@pptx/opc';
import { createMinimalPptx } from '@pptx/testkit';
import { AnimationTimingCodec, installAnimationPlugin } from './index.js';

const SLIDE_URI = '/ppt/slides/slide1.xml';

function slideSource(pkg: OpcPackage): string {
  return new TextDecoder().decode(pkg.requirePart(SLIDE_URI).bytes);
}

function replaceSlideSource(pkg: OpcPackage, source: string): void {
  const part = pkg.requirePart(SLIDE_URI);
  pkg.setPart(part.uri, source, part.contentType);
}

function removeTiming(pkg: OpcPackage): void {
  const xml = LosslessXmlDocument.parse(pkg.requirePart(SLIDE_URI).bytes);
  xml.removeElement(xml.elements('timing')[0]!);
  replaceSlideSource(pkg, xml.serialize());
}

function removePlaybackExtension(pkg: OpcPackage): void {
  const xml = LosslessXmlDocument.parse(pkg.requirePart(SLIDE_URI).bytes);
  const extension = xml.elements('ext').find(
    (candidate) => xml.attribute(candidate, 'uri')?.value
      === '{C13D3E4A-5148-4B6D-A7E7-505054582D4F}',
  );
  xml.removeElement(extension!);
  replaceSlideSource(pkg, xml.serialize());
}

function timingIds(pkg: OpcPackage): readonly number[] {
  const xml = LosslessXmlDocument.parse(pkg.requirePart(SLIDE_URI).bytes);
  return xml.elements('cTn').map((node) => Number(xml.attribute(node, 'id')?.value));
}

describe('AnimationTimingCodec', () => {
  it('adds, decodes, retargets, validates, and removes timing nodes', async () => {
    const pkg = await OpcPackage.open(await createMinimalPptx());
    const codec = installAnimationPlugin({ opcPackage: pkg, codecRegistry: new CodecRegistry() });
    const id = codec.add(SLIDE_URI, {
      effect: 'fade',
      targetShapeId: 2,
      durationMs: 600,
      delayMs: 100,
      trigger: 'on-click',
    });
    expect(codec.tree(SLIDE_URI)).toBeDefined();
    expect(codec.validate(SLIDE_URI)).toEqual([]);
    expect(codec.retargetShape(SLIDE_URI, 2, 2)).toBe(1);
    expect(codec.remove(SLIDE_URI, id)).toBe(true);
  });

  it('reuses native media timing without changing healthy or unsafe imports', async () => {
    const healthyPackage = await OpcPackage.open(await createMinimalPptx());
    const healthyMedia = new MediaCodec(healthyPackage);
    await healthyMedia.addAudio(SLIDE_URI, new Uint8Array([1, 2, 3]), {
      contentType: 'audio/mpeg',
      play: 'auto',
      loop: true,
      volume: 0.5,
    });
    const healthyBefore = slideSource(healthyPackage);
    const healthyCodec = installAnimationPlugin({
      opcPackage: healthyPackage,
      codecRegistry: new CodecRegistry(),
    });
    expect(slideSource(healthyPackage)).toBe(healthyBefore);
    expect(healthyCodec.materializeMediaPlayback(SLIDE_URI)).toBe(0);

    const legacyPackage = await OpcPackage.open(await createMinimalPptx());
    const legacyMedia = new MediaCodec(legacyPackage);
    await legacyMedia.addAudio(SLIDE_URI, Uint8Array.of(1), {
      contentType: 'audio/mpeg',
      play: 'auto',
    });
    await legacyMedia.addVideo(SLIDE_URI, Uint8Array.of(2), {
      contentType: 'video/mp4',
      loop: true,
    });
    removeTiming(legacyPackage);
    const legacyCodec = installAnimationPlugin({
      opcPackage: legacyPackage,
      codecRegistry: new CodecRegistry(),
    });
    expect(LosslessXmlDocument.parse(legacyPackage.requirePart(SLIDE_URI).bytes)
      .elements('cMediaNode')).toHaveLength(2);
    expect(legacyMedia.diagnosticsForSlide(SLIDE_URI, 'powerpoint-current')
      .filter(({ code }) => code.startsWith('MEDIA_TIMING_'))).toEqual([]);
    const legacyOnce = slideSource(legacyPackage);
    expect(legacyCodec.materializeMediaPlayback(SLIDE_URI)).toBe(0);
    expect(slideSource(legacyPackage)).toBe(legacyOnce);
    installAnimationPlugin({
      opcPackage: legacyPackage,
      codecRegistry: new CodecRegistry(),
    });
    expect(slideSource(legacyPackage)).toBe(legacyOnce);

    const stalePackage = await OpcPackage.open(await createMinimalPptx());
    const staleMedia = new MediaCodec(stalePackage);
    await staleMedia.addAudio(SLIDE_URI, Uint8Array.of(3), {
      contentType: 'audio/mpeg',
      volume: 0.5,
    });
    replaceSlideSource(stalePackage, slideSource(stalePackage).replace('vol="50000"', 'vol="25000"'));
    const staleCodec = installAnimationPlugin({
      opcPackage: stalePackage,
      codecRegistry: new CodecRegistry(),
    });
    expect(slideSource(stalePackage)).toContain('vol="50000"');
    expect(staleCodec.materializeMediaPlayback(SLIDE_URI)).toBe(0);

    const nativeOnlyPackage = await OpcPackage.open(await createMinimalPptx());
    const nativeOnlyMedia = new MediaCodec(nativeOnlyPackage);
    await nativeOnlyMedia.addAudio(SLIDE_URI, Uint8Array.of(4), { contentType: 'audio/mpeg' });
    removePlaybackExtension(nativeOnlyPackage);
    const nativeOnlyBefore = slideSource(nativeOnlyPackage);
    installAnimationPlugin({
      opcPackage: nativeOnlyPackage,
      codecRegistry: new CodecRegistry(),
    });
    expect(slideSource(nativeOnlyPackage)).toBe(nativeOnlyBefore);

    const unsupportedPackage = await OpcPackage.open(await createMinimalPptx());
    const unsupportedMedia = new MediaCodec(unsupportedPackage);
    await unsupportedMedia.addAudio(SLIDE_URI, Uint8Array.of(5), {
      contentType: 'audio/mpeg',
      loop: true,
    });
    removePlaybackExtension(unsupportedPackage);
    replaceSlideSource(
      unsupportedPackage,
      slideSource(unsupportedPackage).replace('repeatCount="indefinite"', 'repeatCount="2000"'),
    );
    const unsupportedBefore = slideSource(unsupportedPackage);
    installAnimationPlugin({
      opcPackage: unsupportedPackage,
      codecRegistry: new CodecRegistry(),
    });
    expect(slideSource(unsupportedPackage)).toBe(unsupportedBefore);
    const unsupportedModel = unsupportedMedia.list(SLIDE_URI)[0]!;
    expect(unsupportedMedia.diagnostics(unsupportedModel, 'powerpoint-current').map(({ code }) => code))
      .toContain('MEDIA_TIMING_UNSUPPORTED');
  });

  it('allocates media and general animation IDs above a sparse imported maximum', async () => {
    const pkg = await OpcPackage.open(await createMinimalPptx());
    const media = new MediaCodec(pkg);
    await media.addAudio(SLIDE_URI, Uint8Array.of(1), { contentType: 'audio/mpeg' });
    removeTiming(pkg);
    const codec = new AnimationTimingCodec(pkg);
    codec.add(SLIDE_URI, { effect: 'fade', targetShapeId: 2 });
    replaceSlideSource(
      pkg,
      slideSource(pkg)
        .replace('<p:cTn id="1"', '<p:cTn id="100"')
        .replace('<p:cTn id="2"', '<p:cTn id="500"')
        .replace('<p:cTn id="3"', '<p:cTn id="900"'),
    );

    expect(codec.materializeMediaPlayback(SLIDE_URI)).toBe(1);
    const afterMedia = timingIds(pkg);
    expect(new Set(afterMedia).size).toBe(afterMedia.length);
    expect(Math.min(...afterMedia.filter((id) => id > 900))).toBe(901);
    const previousMaximum = Math.max(...afterMedia);
    const animationId = codec.add(SLIDE_URI, { effect: 'wipe', targetShapeId: 2 });
    expect(animationId).toBe(previousMaximum + 1);
    const finalIds = timingIds(pkg);
    expect(new Set(finalIds).size).toBe(finalIds.length);
  });

  it('materializes one selected direct media picture without touching its peer', async () => {
    const pkg = await OpcPackage.open(await createMinimalPptx());
    const media = new MediaCodec(pkg);
    const audio = await media.addAudio(SLIDE_URI, Uint8Array.of(1), {
      contentType: 'audio/mpeg',
    });
    const video = await media.addVideo(SLIDE_URI, Uint8Array.of(2), {
      contentType: 'video/mp4',
    });
    removeTiming(pkg);

    expect(media.materializePlayback(SLIDE_URI, audio.shapeId)).toBe(1);
    let xml = LosslessXmlDocument.parse(pkg.requirePart(SLIDE_URI).bytes);
    expect(xml.elements('cMediaNode')).toHaveLength(1);
    expect(xml.elements('cMediaNode').flatMap((node) => xml.descendants(node, 'spTgt'))
      .map((target) => Number(xml.attribute(target, 'spid')?.value))).toEqual([audio.shapeId]);
    expect(media.materializePlayback(SLIDE_URI, audio.shapeId)).toBe(0);
    expect(media.materializePlayback(SLIDE_URI, video.shapeId)).toBe(1);
    xml = LosslessXmlDocument.parse(pkg.requirePart(SLIDE_URI).bytes);
    expect(xml.elements('cMediaNode')).toHaveLength(2);
  });
});
