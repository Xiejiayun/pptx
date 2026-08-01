import { describe, expect, it } from 'vitest';
import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
import {
  normalizeImageSourceRectangle,
  readImageSourceRectangle,
  replaceImageSourceRectangle,
  renderImageSourceRectangle,
} from './image-source-rectangle.internal.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';

describe('image source rectangle normalization', () => {
  it('normalizes signed percentages to detached frozen thousandths', () => {
    const input = {
      left: 12.3456,
      top: -20.0004,
      right: 1,
      bottom: -0,
    };
    const normalized = normalizeImageSourceRectangle(input, 'Image source rectangle');
    input.left = 50;
    input.top = 50;

    expect(normalized).toEqual({
      left: 12.346,
      top: -20,
      right: 1,
      bottom: 0,
    });
    expect(Object.isFrozen(normalized)).toBe(true);
  });

  it('accepts null-prototype explicit zero and negative contain values', () => {
    const value = Object.create(null) as Record<string, unknown>;
    value.left = -25;
    value.top = 0;
    value.right = -25;
    value.bottom = 0;

    expect(normalizeImageSourceRectangle(value, 'Image source rectangle')).toEqual({
      left: -25,
      top: 0,
      right: -25,
      bottom: 0,
    });
    expect(normalizeImageSourceRectangle({
      left: 0,
      top: 0,
      right: 0,
      bottom: 0,
    }, 'Image source rectangle')).toEqual({
      left: 0,
      top: 0,
      right: 0,
      bottom: 0,
    });
  });

  it('rejects unsafe containers and descriptors without invoking accessors', () => {
    class Rectangle {
      left = 0;
      top = 0;
      right = 0;
      bottom = 0;
    }
    let reads = 0;
    const accessor = Object.defineProperty({ top: 0, right: 0, bottom: 0 }, 'left', {
      enumerable: true,
      get() {
        reads += 1;
        return 0;
      },
    });
    const inherited = Object.create({ left: 0, top: 0, right: 0, bottom: 0 });
    const symbol = {
      left: 0,
      top: 0,
      right: 0,
      bottom: 0,
      [Symbol('unsafe')]: true,
    };

    for (const value of [null, undefined, [], new Rectangle(), accessor, inherited, symbol]) {
      expect(() => normalizeImageSourceRectangle(value, 'Image source rectangle')).toThrow(
        TypeError,
      );
    }
    expect(reads).toBe(0);
  });

  it('rejects missing, unknown, and invalid edge values', () => {
    const valid = { left: 0, top: 0, right: 0, bottom: 0 };
    const cases = [
      {},
      { top: 0, right: 0, bottom: 0 },
      { ...valid, unknown: true },
      { ...valid, left: '0' },
      { ...valid, left: Number.NaN },
      { ...valid, left: Number.NEGATIVE_INFINITY },
      { ...valid, left: Number.POSITIVE_INFINITY },
      { ...valid, left: -2_147_483.649 },
      { ...valid, left: 100 },
      { ...valid, left: 60, right: 40 },
      { ...valid, top: 75, bottom: 25 },
    ];

    for (const value of cases) {
      expect(() => normalizeImageSourceRectangle(value, 'Image source rectangle')).toThrow();
    }
  });

  it('renders canonical DrawingML integer percentages', () => {
    const value = normalizeImageSourceRectangle({
      left: 12.3456,
      top: -20.0004,
      right: 1,
      bottom: 0,
    }, 'Image source rectangle');

    expect(renderImageSourceRectangle(value)).toBe(
      '<a:srcRect l="12346" t="-20000" r="1000" b="0"/>',
    );
    expect(renderImageSourceRectangle(value, 'd')).toBe(
      '<d:srcRect l="12346" t="-20000" r="1000" b="0"/>',
    );
    expect(renderImageSourceRectangle(value, '')).toBe(
      '<srcRect l="12346" t="-20000" r="1000" b="0"/>',
    );
  });
});

describe('image source rectangle direct state', () => {
  it('reads absent, explicit, defaulted, and alternate-prefix direct state', () => {
    expect(readImageSourceRectangle(...picture('<a:blip/><a:stretch/>'))).toBeUndefined();

    const [explicitXml, explicitPicture] = picture(
      '<a:blip/><a:srcRect l="25000" t="-10000" r="5000" b="0"/><a:stretch/>',
    );
    const first = readImageSourceRectangle(explicitXml, explicitPicture);
    const second = readImageSourceRectangle(explicitXml, explicitPicture);
    expect(first).toEqual({ left: 25, top: -10, right: 5, bottom: 0 });
    expect(second).toEqual(first);
    expect(second).not.toBe(first);
    expect(Object.isFrozen(first)).toBe(true);

    expect(readImageSourceRectangle(...picture(
      '<a:blip/><a:srcRect l="25000"/><a:stretch/>',
    ))).toEqual({ left: 25, top: 0, right: 0, bottom: 0 });
    expect(readImageSourceRectangle(...picture(
      '<d:blip/><d:srcRect l="1000" t="2000" r="3000" b="4000"/><d:stretch/>',
    ))).toEqual({ left: 1, top: 2, right: 3, bottom: 4 });
  });

  it('returns undefined for malformed, duplicate, or namespace-confused direct state', () => {
    const malformed = [
      '<a:srcRect l="+1"/>',
      '<a:srcRect l="01"/>',
      '<a:srcRect l="-0"/>',
      '<a:srcRect l="2147483648"/>',
      '<a:srcRect l="100000"/>',
      '<a:srcRect l="60000" r="40000"/>',
      '<a:srcRect><a:ext/></a:srcRect>',
      '<a:srcRect>unsafe</a:srcRect>',
      '<x:srcRect xmlns:x="urn:wrong" l="0"/>',
      '<a:srcRect/><a:srcRect/>',
    ];
    for (const sourceRectangle of malformed) {
      const [xml, element] = picture(`<a:blip/>${sourceRectangle}<a:stretch/>`);
      const before = xml.serialize();
      expect(readImageSourceRectangle(xml, element), sourceRectangle).toBeUndefined();
      expect(xml.serialize()).toBe(before);
    }
  });

  it('replaces, repairs, clears, and preserves exact no-ops', () => {
    const normalized = normalizeImageSourceRectangle(
      { left: 25, top: -10, right: 5, bottom: 0 },
      'Image source rectangle',
    );
    const [xml, element] = picture(
      '<a:blip/><a:srcRect l="25000" t="-10000" r="5000" b="0"/><a:stretch/>',
    );
    const before = xml.serialize();
    expect(replaceImageSourceRectangle(xml, element, normalized, '/ppt/slides/slide1.xml'))
      .toBe(false);
    expect(xml.serialize()).toBe(before);

    const replacement = normalizeImageSourceRectangle(
      { left: 10, top: 20, right: 30, bottom: 0 },
      'Image source rectangle',
    );
    expect(replaceImageSourceRectangle(xml, element, replacement, '/ppt/slides/slide1.xml'))
      .toBe(true);
    expect(xml.serialize()).toContain(
      '<a:srcRect l="10000" t="20000" r="30000" b="0"/>',
    );

    const [malformedXml, malformedPicture] = picture(
      '<a:blip/><a:srcRect l="bad" custom="remove"/><a:stretch/>',
    );
    expect(replaceImageSourceRectangle(
      malformedXml,
      malformedPicture,
      normalized,
      '/ppt/slides/slide1.xml',
    )).toBe(true);
    expect(malformedXml.serialize()).toContain(
      '<a:srcRect l="25000" t="-10000" r="5000" b="0"/>',
    );
    expect(malformedXml.serialize()).not.toContain('custom="remove"');

    const [clearXml, clearPicture] = picture(
      '<a:blip/><a:srcRect l="bad"/><a:stretch/>',
    );
    expect(replaceImageSourceRectangle(
      clearXml,
      clearPicture,
      undefined,
      '/ppt/slides/slide1.xml',
    )).toBe(true);
    expect(clearXml.serialize()).not.toContain('srcRect');

    const [absentXml, absentPicture] = picture('<a:blip/><a:stretch/>');
    const absentBefore = absentXml.serialize();
    expect(replaceImageSourceRectangle(
      absentXml,
      absentPicture,
      undefined,
      '/ppt/slides/slide1.xml',
    )).toBe(false);
    expect(absentXml.serialize()).toBe(absentBefore);
  });

  it('inserts before the fill choice with a safe local DrawingML binding', () => {
    const [xml, element] = picture(
      '<d:blip xmlns:d="' + DRAWING_NAMESPACE + '" custom="keep">' +
      '<d:extLst><x:keep xmlns:x="urn:keep"/></d:extLst></d:blip>' +
      '<x:neighbor xmlns:x="urn:neighbor"/>' +
      '<d:stretch xmlns:d="' + DRAWING_NAMESPACE + '" custom="stretch">' +
      '<d:fillRect/></d:stretch><x:tail xmlns:x="urn:tail"/>',
      false,
    );
    const normalized = normalizeImageSourceRectangle(
      { left: 25, top: -10, right: 5, bottom: 0 },
      'Image source rectangle',
    );
    expect(replaceImageSourceRectangle(xml, element, normalized, '/ppt/slides/slide1.xml'))
      .toBe(true);
    const source = xml.serialize();
    expect(source).toContain(
      '<d:srcRect xmlns:d="' + DRAWING_NAMESPACE +
      '" l="25000" t="-10000" r="5000" b="0"/>',
    );
    expect(source.indexOf('<x:neighbor')).toBeLessThan(source.indexOf('<d:srcRect'));
    expect(source.indexOf('<d:srcRect')).toBeLessThan(source.indexOf('<d:stretch'));
    expect(source).toContain('<x:keep xmlns:x="urn:keep"/>');
    expect(source).toContain('custom="stretch"');
    expect(source).toContain('<x:tail xmlns:x="urn:tail"/>');
  });

  it('rejects ambiguous or unsafe ownership without mutation', () => {
    const normalized = normalizeImageSourceRectangle(
      { left: 0, top: 0, right: 0, bottom: 0 },
      'Image source rectangle',
    );
    for (const [xml, element] of [
      picture('<a:blip/><a:srcRect/><a:srcRect/><a:stretch/>'),
      picture('<a:blip/><x:srcRect xmlns:x="urn:wrong"/><a:stretch/>'),
      picture('<a:blip/><a:stretch/>', true, true),
    ]) {
      const before = xml.serialize();
      expect(() => replaceImageSourceRectangle(
        xml,
        element,
        normalized,
        '/ppt/slides/slide1.xml',
      )).toThrow(ModelParseError);
      expect(xml.serialize()).toBe(before);
    }
  });
});

function picture(
  fill: string,
  drawingOnRoot = true,
  duplicateFill = false,
): [LosslessXmlDocument, XmlElement] {
  const drawing = drawingOnRoot
    ? ` xmlns:a="${DRAWING_NAMESPACE}" xmlns:d="${DRAWING_NAMESPACE}"`
    : '';
  const fills = duplicateFill
    ? `<p:blipFill>${fill}</p:blipFill><p:blipFill><a:blip/></p:blipFill>`
    : `<p:blipFill>${fill}</p:blipFill>`;
  const xml = LosslessXmlDocument.parse(
    `<p:pic xmlns:p="${PRESENTATION_NAMESPACE}"${drawing}>${fills}</p:pic>`,
  );
  return [xml, xml.roots[0]!];
}
