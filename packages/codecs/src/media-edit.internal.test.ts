import { describe, expect, it } from 'vitest';
import { LosslessXmlDocument } from '@pptx/lossless-xml';
import {
  normalizeMediaAltText,
  normalizeMediaName,
  normalizeMediaPlaybackSettings,
  replaceMediaMetadataAttribute,
  replaceMediaPlaybackExtension,
} from './media-edit.internal.js';

describe('media editing normalization and patches', () => {
  it('normalizes strict detached playback objects without invoking accessors', () => {
    const value = Object.create(null) as Record<string, unknown>;
    value.play = 'auto';
    value.loop = true;
    value.volume = 0.25;
    const normalized = normalizeMediaPlaybackSettings(value)!;
    value.play = 'click';
    expect(normalized).toEqual({
      play: 'auto',
      loop: true,
      hideWhenStopped: false,
      volume: 0.25,
    });
    expect(Object.isFrozen(normalized)).toBe(true);

    let reads = 0;
    const accessor = Object.defineProperty({}, 'play', {
      enumerable: true,
      get() { reads += 1; return 'auto'; },
    });
    expect(() => normalizeMediaPlaybackSettings(accessor)).toThrow(/data property/);
    expect(reads).toBe(0);
    for (const invalid of [null, [], new Date(), { extra: true }, { volume: 2 }, { loop: 1 }]) {
      expect(() => normalizeMediaPlaybackSettings(invalid)).toThrow();
    }
  });

  it('validates metadata and patches only owned attributes', () => {
    expect(normalizeMediaName('Audio & <one>')).toBe('Audio & <one>');
    expect(normalizeMediaAltText(undefined)).toBeUndefined();
    expect(() => normalizeMediaName(undefined)).toThrow();
    expect(() => normalizeMediaAltText('\u0001')).toThrow(/invalid XML/);
    const xml = LosslessXmlDocument.parse(
      '<p:pic xmlns:p="p"><p:nvPicPr><p:cNvPr id="2" name="Old" keep="1"/>' +
      '<p:nvPr/></p:nvPicPr><x:keep xmlns:x="x"/></p:pic>',
    );
    const picture = xml.elements('pic')[0]!;
    expect(replaceMediaMetadataAttribute(xml, picture, 'name', 'Audio & <one>')).toBe(true);
    expect(replaceMediaMetadataAttribute(xml, picture, 'descr', '')).toBe(true);
    expect(xml.serialize()).toContain('name="Audio &amp; &lt;one&gt;" keep="1" descr=""');
    expect(xml.serialize()).toContain('<x:keep xmlns:x="x"/>');
  });

  it('adds, replaces, clears, and no-ops the owned playback extension', () => {
    const input = '<p:pic xmlns:p="p"><p:nvPicPr><p:cNvPr id="2" name="Media"/><p:nvPr>' +
      '<a:audioFile xmlns:a="a"/></p:nvPr></p:nvPicPr><p:timing keep="1"/></p:pic>';
    const settings = normalizeMediaPlaybackSettings({ play: 'auto', loop: true, volume: 0.5 })!;
    const first = LosslessXmlDocument.parse(input);
    expect(replaceMediaPlaybackExtension(first, first.elements('pic')[0]!, settings)).toBe(true);
    const rendered = first.serialize();
    expect(rendered).toContain('play="auto" loop="1" hideWhenStopped="0" volume="50000"');
    expect(rendered).toContain('<p:timing keep="1"/>');

    const second = LosslessXmlDocument.parse(rendered);
    expect(replaceMediaPlaybackExtension(second, second.elements('pic')[0]!, settings)).toBe(false);
    expect(second.serialize()).toBe(rendered);
    expect(replaceMediaPlaybackExtension(second, second.elements('pic')[0]!, undefined)).toBe(true);
    expect(second.serialize()).not.toContain('playback');
    expect(second.serialize()).toContain('<p:timing keep="1"/>');
  });
});
