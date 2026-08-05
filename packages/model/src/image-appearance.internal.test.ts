import { describe, expect, it } from 'vitest';
import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import {
  normalizeImageAltText,
  normalizeImageRounding,
  normalizeImageTransparency,
  normalizeShapeName,
  readImageAltText,
  readImageRounding,
  readImageTransparency,
  replaceImageAltText,
  replaceImageRounding,
  replaceImageTransparency,
  replaceShapeName,
} from './image-appearance.internal.js';
import { readShapeShadow, replaceShapeShadow } from './shape-shadow.internal.js';
import { normalizeShapeShadow } from './simple-shadow.internal.js';

const PART_URI = '/ppt/slides/slide1.xml';
const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function fixture(
  geometry = '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>',
  blipChildren = '<a:lum bright="1000"/><a:extLst><a:ext uri="keep"/></a:extLst>',
  effects = '<a:effectLst><a:glow rad="1"/></a:effectLst>',
  nonVisual = '<p:cNvPr id="7" name="Before" descr="Before alt" data-keep="yes"/>',
): string {
  return `<p:pic xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}" data-owner="keep">`
    + `<p:nvPicPr>${nonVisual}<p:cNvPicPr/><p:nvPr/></p:nvPicPr>`
    + `<p:blipFill><a:blip r:embed="rId1" data-blip="keep">${blipChildren}</a:blip>`
    + '<a:stretch><a:fillRect/></a:stretch></p:blipFill>'
    + `<p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm>`
    + `${geometry}<a:ln w="9"/>${effects}<a:extLst><a:ext uri="shape-keep"/></a:extLst>`
    + '</p:spPr></p:pic>';
}

function parse(source: string): { xml: LosslessXmlDocument; picture: XmlElement } {
  const xml = LosslessXmlDocument.parse(source);
  return { xml, picture: xml.roots[0]! };
}

describe('image appearance normalization', () => {
  it('accepts strict legal values and canonicalizes fractional transparency', () => {
    expect(normalizeShapeName('')).toBe('');
    expect(normalizeShapeName('Shape & <name>')).toBe('Shape & <name>');
    expect(normalizeImageAltText(undefined)).toBeUndefined();
    expect(normalizeImageAltText('')).toBe('');
    expect(normalizeImageRounding(true)).toBe(true);
    expect(normalizeImageTransparency(-0)).toBe(0);
    expect(normalizeImageTransparency(25.0004)).toBe(25);
    expect(normalizeImageTransparency(100)).toBe(100);
  });

  it('rejects invalid strings, booleans, and percentages', () => {
    for (const value of [undefined, null, 1, 'bad\u0000name']) {
      expect(() => normalizeShapeName(value)).toThrow('Shape name');
    }
    for (const value of [null, 1, 'bad\u000Btext']) {
      expect(() => normalizeImageAltText(value)).toThrow(TypeError);
    }
    for (const value of [undefined, 0, 'true']) {
      expect(() => normalizeImageRounding(value)).toThrow(TypeError);
    }
    for (const value of [undefined, Number.NaN, Number.POSITIVE_INFINITY, -0.001, 100.001]) {
      expect(() => normalizeImageTransparency(value)).toThrow();
    }
  });
});

describe('lossless image appearance editing', () => {
  it('reads defaults and edits each owned fragment without disturbing neighbors', () => {
    const initial = parse(fixture());
    expect(readImageAltText(initial.xml, initial.picture, PART_URI)).toBe('Before alt');
    expect(readImageRounding(initial.picture)).toBe(false);
    expect(readImageTransparency(initial.picture)).toBe(0);
    expect(readShapeShadow(initial.xml, initial.picture)).toBeUndefined();

    const named = parse(fixture());
    expect(replaceShapeName(named.xml, named.picture, 'After & <name>', PART_URI)).toBe(true);
    const namedSource = named.xml.serialize();
    expect(namedSource).toContain('name="After &amp; &lt;name&gt;"');
    expect(namedSource).toContain('descr="Before alt" data-keep="yes"');
    expect(namedSource).toContain('r:embed="rId1" data-blip="keep"');

    const alt = parse(fixture());
    expect(replaceImageAltText(alt.xml, alt.picture, undefined, PART_URI)).toBe(true);
    expect(alt.xml.serialize()).not.toContain(' descr=');
    expect(alt.xml.serialize()).toContain('name="Before" data-keep="yes"');

    const rounded = parse(fixture());
    expect(replaceImageRounding(rounded.xml, rounded.picture, true, PART_URI)).toBe(true);
    expect(rounded.xml.serialize()).toContain(
      '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:ln w="9"/>',
    );

    const transparent = parse(fixture());
    expect(replaceImageTransparency(transparent.xml, transparent.picture, 25, PART_URI)).toBe(true);
    const transparentSource = transparent.xml.serialize();
    expect(transparentSource).toContain(
      '<a:lum bright="1000"/><a:alphaModFix amt="75000"/><a:extLst>',
    );
    const reopenedTransparency = parse(transparentSource);
    expect(readImageTransparency(reopenedTransparency.picture)).toBe(25);
    expect(replaceImageTransparency(
      reopenedTransparency.xml,
      reopenedTransparency.picture,
      0,
      PART_URI,
    )).toBe(true);
    expect(reopenedTransparency.xml.serialize()).not.toContain('alphaModFix');

    const explicitZero = parse(fixture(
      undefined,
      '<a:alphaModFix amt="100000"/><a:extLst><a:ext uri="keep"/></a:extLst>',
    ));
    expect(readImageTransparency(explicitZero.picture)).toBe(0);
    expect(replaceImageTransparency(
      explicitZero.xml,
      explicitZero.picture,
      0,
      PART_URI,
    )).toBe(true);
    expect(explicitZero.xml.serialize()).not.toContain('alphaModFix');

    const shadowed = parse(fixture());
    const shadow = normalizeShapeShadow({
      kind: 'outer',
      color: { kind: 'srgb', value: '123456' },
      opacity: 0.5,
      blur: 3,
      angle: 30,
      distance: 2,
      rotateWithShape: true,
    }, 'Image shadow');
    expect(replaceShapeShadow(shadowed.xml, shadowed.picture, shadow, PART_URI)).toBe(true);
    const shadowedSource = shadowed.xml.serialize();
    expect(shadowedSource).toContain('<a:glow rad="1"/><a:outerShdw');
    expect(shadowedSource).toContain('rotWithShape="1"');
    const reopenedShadow = parse(shadowedSource);
    expect(readShapeShadow(reopenedShadow.xml, reopenedShadow.picture)).toEqual(shadow);
    expect(shadowedSource).toContain('<a:ext uri="shape-keep"/>');
  });

  it('keeps semantic no-ops byte-identical, including explicit empty metadata', () => {
    const source = fixture(
      '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>',
      '<a:alphaModFix amt="75000"/><a:extLst><a:ext uri="keep"/></a:extLst>',
      '<a:effectLst/>',
      '<p:cNvPr id="7" name="" descr=""/>',
    );
    for (const edit of [
      (xml: LosslessXmlDocument, picture: XmlElement) =>
        replaceShapeName(xml, picture, '', PART_URI),
      (xml: LosslessXmlDocument, picture: XmlElement) =>
        replaceImageAltText(xml, picture, '', PART_URI),
      (xml: LosslessXmlDocument, picture: XmlElement) =>
        replaceImageRounding(xml, picture, true, PART_URI),
      (xml: LosslessXmlDocument, picture: XmlElement) =>
        replaceImageTransparency(xml, picture, 25, PART_URI),
    ]) {
      const { xml, picture } = parse(source);
      expect(edit(xml, picture)).toBe(false);
      expect(xml.serialize()).toBe(source);
    }
  });

  it('supports alternate presentation and drawing prefixes', () => {
    const source = fixture()
      .replaceAll('<p:', '<q:')
      .replaceAll('</p:', '</q:')
      .replace('xmlns:p=', 'xmlns:q=')
      .replaceAll('<a:', '<d:')
      .replaceAll('</a:', '</d:')
      .replace('xmlns:a=', 'xmlns:d=');
    const { xml, picture } = parse(source);
    expect(replaceImageRounding(xml, picture, true, PART_URI)).toBe(true);
    const rounded = parse(xml.serialize());
    expect(replaceImageTransparency(rounded.xml, rounded.picture, 100, PART_URI)).toBe(true);
    expect(rounded.xml.serialize()).toContain(
      '<d:alphaModFix amt="0"/><d:extLst>',
    );
  });

  it('rejects ambiguous identity, geometry, and blip ownership', () => {
    const cases = [
      fixture(undefined, undefined, undefined,
        '<p:cNvPr id="7" name="Before" name="Again"/>'),
      fixture(
        '<a:prstGeom prst="rect"/><a:prstGeom prst="ellipse"/>',
      ),
      fixture(undefined,
        '<a:alphaModFix amt="75000"/><a:alphaModFix amt="50000"/>'),
      fixture(undefined, '<x:custom xmlns:x="urn:unknown"/>'),
    ] as const;

    const duplicateName = parse(cases[0]);
    expect(() => replaceShapeName(
      duplicateName.xml,
      duplicateName.picture,
      'After',
      PART_URI,
    )).toThrow();

    const duplicateGeometry = parse(cases[1]);
    expect(readImageRounding(duplicateGeometry.picture)).toBeUndefined();
    expect(() => replaceImageRounding(
      duplicateGeometry.xml,
      duplicateGeometry.picture,
      true,
      PART_URI,
    )).toThrow();

    for (const source of cases.slice(2)) {
      const { xml, picture } = parse(source);
      expect(readImageTransparency(picture)).toBeUndefined();
      expect(() => replaceImageTransparency(xml, picture, 25, PART_URI)).toThrow();
    }
  });

  it('keeps image-only editors away from non-picture owners', () => {
    const source = fixture().replace('<p:pic ', '<p:sp ').replace('</p:pic>', '</p:sp>')
      .replace('<p:nvPicPr>', '<p:nvSpPr>')
      .replace('</p:nvPicPr>', '</p:nvSpPr>');
    const { xml, picture: shape } = parse(source);
    expect(readImageAltText(xml, shape, PART_URI)).toBeUndefined();
    expect(() => replaceImageAltText(xml, shape, 'After', PART_URI)).toThrow();
    expect(xml.serialize()).toBe(source);
  });

  it('rejects same-local-name owners from conflicting namespaces', () => {
    const sources = [
      fixture().replace(
        'data-owner="keep">',
        'xmlns:x="urn:wrong" data-owner="keep"><x:nvPicPr/>',
      ),
      fixture().replace(
        '<p:cNvPr id="7"',
        '<x:cNvPr xmlns:x="urn:wrong"/><p:cNvPr id="7"',
      ),
      fixture().replace(
        'data-owner="keep">',
        'xmlns:x="urn:wrong" data-owner="keep"><x:blipFill/>',
      ),
      fixture().replace(
        'data-owner="keep">',
        'xmlns:x="urn:wrong" data-owner="keep"><x:spPr/>',
      ),
    ] as const;

    const identityOwner = parse(sources[0]);
    expect(readImageAltText(identityOwner.xml, identityOwner.picture, PART_URI)).toBeUndefined();
    expect(() => replaceImageAltText(
      identityOwner.xml,
      identityOwner.picture,
      'After',
      PART_URI,
    )).toThrow();
    expect(identityOwner.xml.serialize()).toBe(sources[0]);

    const identityProperties = parse(sources[1]);
    expect(() => replaceShapeName(
      identityProperties.xml,
      identityProperties.picture,
      'After',
      PART_URI,
    )).toThrow();
    expect(identityProperties.xml.serialize()).toBe(sources[1]);

    const transparencyOwner = parse(sources[2]);
    expect(readImageTransparency(transparencyOwner.picture)).toBeUndefined();
    expect(() => replaceImageTransparency(
      transparencyOwner.xml,
      transparencyOwner.picture,
      25,
      PART_URI,
    )).toThrow();
    expect(transparencyOwner.xml.serialize()).toBe(sources[2]);

    const roundingOwner = parse(sources[3]);
    expect(readImageRounding(roundingOwner.picture)).toBeUndefined();
    expect(() => replaceImageRounding(
      roundingOwner.xml,
      roundingOwner.picture,
      true,
      PART_URI,
    )).toThrow();
    expect(roundingOwner.xml.serialize()).toBe(sources[3]);
  });
});
