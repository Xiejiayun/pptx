import { describe, expect, it } from 'vitest';
import { CodecRegistry } from '@pptx/codecs';
import { OpcPackage } from '@pptx/opc';
import { createMinimalPptx } from '@pptx/testkit';
import { installTransitionPlugin } from './index.js';

describe('TransitionCodec', () => {
  it('adds, edits, reads, diagnoses, and clears transitions', async () => {
    const pkg = await OpcPackage.open(await createMinimalPptx());
    const codec = installTransitionPlugin({ opcPackage: pkg, codecRegistry: new CodecRegistry() });
    codec.set('/ppt/slides/slide1.xml', {
      effect: 'fade',
      speed: 'fast',
      durationMs: 750,
      advanceOnClick: false,
      advanceAfterMs: 2_000,
    });
    expect(codec.get('/ppt/slides/slide1.xml')).toMatchObject({
      effect: 'fade',
      speed: 'fast',
      durationMs: 750,
      advanceOnClick: false,
      advanceAfterMs: 2_000,
    });
    expect(codec.diagnostics({ effect: 'morph' }, '/ppt/slides/slide1.xml', 'powerpoint-2010')[0]?.code).toBe(
      'TRANSITION_MORPH_PRESERVED_ONLY',
    );
    pkg.setPart('/ppt/media/sound1.wav', new Uint8Array([1, 2, 3]), 'audio/wav');
    const soundRelationshipId = codec.setSound('/ppt/slides/slide1.xml', '/ppt/media/sound1.wav', true);
    expect(codec.get('/ppt/slides/slide1.xml')?.soundRelationshipId).toBe(soundRelationshipId);
    expect(codec.clearSound('/ppt/slides/slide1.xml')).toBe(true);
    codec.clear('/ppt/slides/slide1.xml');
    expect(codec.get('/ppt/slides/slide1.xml')).toBeUndefined();
  });
});
