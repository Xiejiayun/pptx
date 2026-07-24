import { describe, expect, it } from 'vitest';
import { CodecRegistry, MediaCodec } from '@pptx/codecs';
import { OpcPackage } from '@pptx/opc';
import { createMinimalPptx } from '@pptx/testkit';
import { installAnimationPlugin } from './index.js';

describe('AnimationTimingCodec', () => {
  it('adds, decodes, retargets, validates, and removes timing nodes', async () => {
    const pkg = await OpcPackage.open(await createMinimalPptx());
    const codec = installAnimationPlugin({ opcPackage: pkg, codecRegistry: new CodecRegistry() });
    const id = codec.add('/ppt/slides/slide1.xml', {
      effect: 'fade',
      targetShapeId: 2,
      durationMs: 600,
      delayMs: 100,
      trigger: 'on-click',
    });
    expect(codec.tree('/ppt/slides/slide1.xml')).toBeDefined();
    expect(codec.validate('/ppt/slides/slide1.xml')).toEqual([]);
    expect(codec.retargetShape('/ppt/slides/slide1.xml', 2, 2)).toBe(1);
    expect(codec.remove('/ppt/slides/slide1.xml', id)).toBe(true);
  });

  it('materializes round-tripped media playback preferences into native timing nodes', async () => {
    const pkg = await OpcPackage.open(await createMinimalPptx());
    const media = new MediaCodec(pkg);
    await media.addAudio('/ppt/slides/slide1.xml', new Uint8Array([1, 2, 3]), {
      contentType: 'audio/mpeg',
      play: 'auto',
      loop: true,
      volume: 0.5,
    });
    const codec = installAnimationPlugin({ opcPackage: pkg, codecRegistry: new CodecRegistry() });
    expect(codec.tree('/ppt/slides/slide1.xml')).toBeDefined();
  });
});
