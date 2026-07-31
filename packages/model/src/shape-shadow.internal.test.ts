import { LosslessXmlDocument } from '@pptx/lossless-xml';
import { describe, expect, it } from 'vitest';
import { ModelParseError } from './errors.js';
import { normalizeShapeShadow } from './simple-shadow.internal.js';
import {
  readShapeShadow,
  replaceShapeShadow,
} from './shape-shadow.internal.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const PART_URI = '/ppt/slides/slide1.xml';

function parse(source: string) {
  const xml = LosslessXmlDocument.parse(source);
  const shape = xml.roots[0];
  if (!shape) throw new Error('Fixture has no shape');
  return { xml, shape };
}

function properties(effect = '', before = '', after = ''): string {
  return '<p:spPr keep="PROPERTIES"><a:xfrm><a:off x="1" y="2"/>' +
    '<a:ext cx="3" cy="4"/></a:xfrm>' +
    '<a:prstGeom prst="rect"><a:avLst><a:gd name="adj" fmla="val 1"/>' +
    `</a:avLst></a:prstGeom>${before}<a:solidFill>` +
    '<a:srgbClr val="EEEEEE"/></a:solidFill>' +
    '<a:ln w="12700"><a:solidFill><a:srgbClr val="111111"/></a:solidFill>' +
    '<a:prstDash val="dash"/><a:headEnd type="triangle"/>' +
    `<a:tailEnd type="arrow"/></a:ln>${effect}${after}` +
    '<a:scene3d/><a:sp3d/><a:extLst><a:ext uri="urn:keep">' +
    '<x:opaque xmlns:x="urn:test">KEEP</x:opaque></a:ext></a:extLst></p:spPr>';
}

function fixture(
  shapeProperties: string,
  options: { readonly root?: string; readonly presentationNamespace?: string } = {},
): string {
  const root = options.root ?? 'p:sp';
  const namespace = options.presentationNamespace ?? PRESENTATION_NAMESPACE;
  return `<${root} xmlns:p="${namespace}" xmlns:a="${DRAWING_NAMESPACE}" ` +
    'xmlns:r="urn:r"><p:nvSpPr><p:cNvPr id="7" name="Keep">' +
    '<a:hlinkClick r:id="rId9"/></p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
    `${shapeProperties}<p:txBody><a:bodyPr/><a:p><a:r><a:t>KEEP</a:t>` +
    `</a:r></a:p></p:txBody></${root}>`;
}

const OUTER = '<a:outerShdw sx="100000" sy="100000" kx="0" ky="0" ' +
  'algn="bl" rotWithShape="0" blurRad="101600" dist="50800" dir="16200000">' +
  '<a:srgbClr val="000000"><a:alpha val="75000"/></a:srgbClr></a:outerShdw>';
const INNER = '<a:innerShdw blurRad="0" dist="0" dir="0">' +
  '<a:schemeClr val="accent2"><a:alpha val="0"/></a:schemeClr></a:innerShdw>';

describe('shape shadow inspection', () => {
  it('reads unique direct outer and inner shadows without mutation', () => {
    const sources = [
      fixture(properties(`<a:effectLst>${OUTER}</a:effectLst>`)),
      fixture(properties(`<a:effectLst>${INNER}</a:effectLst>`)),
      `<q:sp xmlns:q="${PRESENTATION_NAMESPACE}" xmlns:d="${DRAWING_NAMESPACE}">` +
        '<q:spPr><d:prstGeom prst="ellipse"><d:avLst/></d:prstGeom>' +
        '<d:effectLst><d:outerShdw><d:schemeClr val="accent4"/>' +
        '</d:outerShdw></d:effectLst></q:spPr></q:sp>',
    ];
    const expected = [
      {
        kind: 'outer',
        color: { kind: 'srgb', value: '000000' },
        opacity: 0.75,
        blur: 8,
        angle: 270,
        distance: 4,
        rotateWithShape: false,
      },
      {
        kind: 'inner',
        color: { kind: 'scheme', value: 'accent2' },
        opacity: 0,
        blur: 0,
        angle: 0,
        distance: 0,
      },
      {
        kind: 'outer',
        color: { kind: 'scheme', value: 'accent4' },
        opacity: 1,
        blur: 0,
        angle: 0,
        distance: 0,
        rotateWithShape: true,
      },
    ];

    for (let index = 0; index < sources.length; index += 1) {
      const source = sources[index]!;
      const { xml, shape } = parse(source);
      const shadow = readShapeShadow(xml, shape);
      expect(shadow).toEqual(expected[index]);
      expect(Object.isFrozen(shadow)).toBe(true);
      expect(Object.isFrozen(shadow?.color)).toBe(true);
      expect(xml.changed).toBe(false);
      expect(xml.serialize()).toBe(source);
    }
  });

  it('returns undefined for absent, ambiguous, malformed, namespace-unsafe, or unordered state', () => {
    const sources = [
      fixture(properties()),
      fixture(properties('<a:effectLst/>')),
      fixture(properties(`<a:effectDag>${OUTER}</a:effectDag>`)),
      fixture(properties('<a:effectLst/><a:effectLst/>')),
      fixture(properties(`<x:effectLst xmlns:x="urn:wrong">${OUTER}</x:effectLst>`)),
      fixture(properties(`<a:effectLst>${OUTER}${INNER}</a:effectLst>`)),
      fixture(properties(`<a:effectLst>${OUTER}${OUTER}</a:effectLst>`)),
      fixture(properties('<a:effectLst><a:reflection/><a:glow/></a:effectLst>')),
      fixture(properties('<a:effectLst><x:glow xmlns:x="urn:wrong"/></a:effectLst>')),
      fixture(properties('<a:effectLst custom="1"/>')),
      fixture(properties('<a:effectLst>TEXT</a:effectLst>')),
      fixture(properties('<a:effectLst><a:outerShdw blurRad="bad">' +
        '<a:srgbClr val="000000"/></a:outerShdw></a:effectLst>')),
      fixture(properties('<a:effectLst><a:outerShdw sx="100000">' +
        '<a:srgbClr val="000000"/></a:outerShdw></a:effectLst>')),
      fixture(properties('<a:effectLst><a:outerShdw><a:srgbClr val="000000">' +
        '<a:alpha val="100001"/></a:srgbClr></a:outerShdw></a:effectLst>')),
      fixture(properties('<a:effectLst><a:outerShdw>' +
        '<x:srgbClr xmlns:x="urn:wrong" val="000000"/>' +
        '</a:outerShdw></a:effectLst>')),
      fixture(properties(OUTER), { root: 'p:pic' }),
      fixture(properties(OUTER), { presentationNamespace: 'urn:wrong' }),
      fixture('<p:spPr/><p:spPr/>'),
      fixture('<x:spPr xmlns:x="urn:wrong"/>'),
    ];
    for (const source of sources) {
      const { xml, shape } = parse(source);
      expect(readShapeShadow(xml, shape), source).toBeUndefined();
      expect(xml.changed, source).toBe(false);
      expect(xml.serialize(), source).toBe(source);
    }
  });
});

describe('shape shadow replacement', () => {
  it('inserts into missing, self-closing, and expanded effect lists at every schema stage', () => {
    const siblings =
      '<a:blur/><a:fillOverlay/><a:glow/>' +
      '<a:prstShdw/><a:reflection/><a:softEdge/>';
    const cases = [
      fixture(properties()),
      fixture(properties('<a:effectLst/>')),
      fixture(properties('<a:effectLst></a:effectLst>')),
      fixture(properties(`<a:effectLst>${siblings}</a:effectLst>`)),
      `<q:sp xmlns:q="${PRESENTATION_NAMESPACE}" xmlns:d="${DRAWING_NAMESPACE}">` +
        '<q:spPr><d:prstGeom prst="rect"><d:avLst/></d:prstGeom>' +
        '<d:ln/><d:scene3d/></q:spPr></q:sp>',
    ];
    for (const source of cases) {
      const { xml, shape } = parse(source);
      expect(replaceShapeShadow(
        xml,
        shape,
        normalizeShapeShadow({ kind: 'outer' }, 'Shape shadow'),
        PART_URI,
      ), source).toBe(true);
      const updated = xml.serialize();
      expect(updated.match(/effectLst/g), source).toHaveLength(2);
      expect(updated.match(/outerShdw/g), source).toHaveLength(2);
      expect(updated, source).toContain('blurRad="101600" dist="50800" dir="16200000"');
      expect(updated.indexOf('outerShdw'), source).toBeLessThan(
        updated.indexOf('prstShdw') < 0 ? Number.MAX_SAFE_INTEGER : updated.indexOf('prstShdw'),
      );
      expect(updated.indexOf('outerShdw'), source).toBeLessThan(
        updated.indexOf('reflection') < 0 ? Number.MAX_SAFE_INTEGER : updated.indexOf('reflection'),
      );
      if (source.includes('KEEP')) expect(updated, source).toContain('KEEP');
      if (source.includes('r:id="rId9"')) expect(updated, source).toContain('r:id="rId9"');
    }

    const source = fixture(properties(
      '<a:effectLst><a:glow/><a:outerShdw><a:srgbClr val="111111"/>' +
      '</a:outerShdw><a:prstShdw/><a:reflection/></a:effectLst>',
    ));
    const { xml, shape } = parse(source);
    expect(replaceShapeShadow(
      xml,
      shape,
      normalizeShapeShadow({ kind: 'inner', color: { kind: 'scheme', value: 'accent3' } }, 'x'),
      PART_URI,
    )).toBe(true);
    const updated = xml.serialize();
    expect(updated.indexOf('<a:glow/>')).toBeLessThan(updated.indexOf('<a:innerShdw'));
    expect(updated.indexOf('<a:innerShdw')).toBeLessThan(updated.indexOf('<a:prstShdw'));
    expect(updated).not.toContain('<a:outerShdw');
  });

  it('keeps same-value assignments byte-exact and patches same-kind fields locally', () => {
    const source = `<q:sp xmlns:q="${PRESENTATION_NAMESPACE}" ` +
      `xmlns:d="${DRAWING_NAMESPACE}"><q:spPr><d:prstGeom prst="rect">` +
      '<d:avLst/></d:prstGeom><d:effectLst xmlns:k="urn:keep"><d:glow/>' +
      '<d:outerShdw sx="100000" sy="100000" kx="0" ky="0" algn="bl" ' +
      "rotWithShape='true' blurRad='12700' dist='25400' dir='1800000'>" +
      "<d:srgbClr val='112233'><d:alpha val='50000'/></d:srgbClr>" +
      '</d:outerShdw><d:reflection/></d:effectLst></q:spPr></q:sp>';
    const current = normalizeShapeShadow({
      kind: 'outer',
      color: { kind: 'srgb', value: '112233' },
      opacity: 0.5,
      blur: 1,
      angle: 30,
      distance: 2,
      rotateWithShape: true,
    }, 'x');
    const same = parse(source);
    expect(replaceShapeShadow(same.xml, same.shape, current, PART_URI)).toBe(false);
    expect(same.xml.changed).toBe(false);
    expect(same.xml.serialize()).toBe(source);

    const edits = [
      [{ ...current, blur: 2 }, "blurRad='12700'", "blurRad='25400'"],
      [{ ...current, distance: 3 }, "dist='25400'", "dist='38100'"],
      [{ ...current, angle: 45 }, "dir='1800000'", "dir='2700000'"],
      [{ ...current, rotateWithShape: false }, "rotWithShape='true'", "rotWithShape='0'"],
      [
        { ...current, color: { kind: 'srgb' as const, value: 'ABCDEF' } },
        "val='112233'",
        "val='ABCDEF'",
      ],
      [{ ...current, opacity: 0.25 }, "val='50000'", "val='25000'"],
    ] as const;
    for (const [target, before, after] of edits) {
      const parsed = parse(source);
      expect(replaceShapeShadow(parsed.xml, parsed.shape, target, PART_URI)).toBe(true);
      const updated = parsed.xml.serialize();
      expect(updated, after).toBe(source.replace(before, after));
    }
  });

  it('adds omitted attributes and alpha while preserving explicit zero and lexical neighbors', () => {
    const source = `<q:sp xmlns:q="${PRESENTATION_NAMESPACE}" ` +
      `xmlns:d="${DRAWING_NAMESPACE}"><q:spPr><d:prstGeom prst="rect"/>` +
      '<d:effectLst><d:outerShdw sx="100000" sy="100000" kx="0" ky="0" ' +
      'algn="bl"><d:srgbClr val="000000"/></d:outerShdw></d:effectLst>' +
      '</q:spPr></q:sp>';
    const { xml, shape } = parse(source);
    expect(replaceShapeShadow(xml, shape, normalizeShapeShadow({
      kind: 'outer',
      opacity: 0,
      blur: 0,
      angle: 0,
      distance: 0,
      rotateWithShape: false,
    }, 'x'), PART_URI)).toBe(true);
    const updated = xml.serialize();
    expect(updated).toContain('rotWithShape="0"');
    expect(updated).not.toContain('blurRad=');
    expect(updated).not.toContain(' dist=');
    expect(updated).not.toContain(' dir=');
    expect(updated).toContain('<d:srgbClr val="000000"><d:alpha val="0"/>' +
      '</d:srgbClr>');

    const geometry = parse(source);
    expect(replaceShapeShadow(geometry.xml, geometry.shape, normalizeShapeShadow({
      kind: 'outer',
      opacity: 1,
      blur: 2,
      angle: 45,
      distance: 3,
      rotateWithShape: true,
    }, 'x'), PART_URI)).toBe(true);
    const geometryUpdated = geometry.xml.serialize();
    expect(geometryUpdated).toContain(
      'blurRad="25400" dist="38100" dir="2700000"',
    );
    expect(geometryUpdated).not.toContain('rotWithShape=');
    expect(geometryUpdated).toContain('<d:srgbClr val="000000"/>');
  });

  it('switches color and shadow kinds, clears only the shadow, and preserves effect siblings', () => {
    const source = fixture(properties(
      '<a:effectLst xmlns:k="urn:keep"><a:blur/><a:glow keep="GLOW"/>' +
      `${OUTER}<a:prstShdw keep="PRESET"/><a:reflection keep="REFLECTION"/>` +
      '<a:softEdge keep="SOFT"/></a:effectLst>',
    ));
    const color = parse(source);
    expect(replaceShapeShadow(color.xml, color.shape, normalizeShapeShadow({
      kind: 'outer',
      color: { kind: 'scheme', value: 'accent5' },
    }, 'x'), PART_URI)).toBe(true);
    const colorUpdated = color.xml.serialize();
    expect(colorUpdated).toContain('<a:schemeClr val="accent5"><a:alpha val="75000"/>' +
      '</a:schemeClr>');
    expect(colorUpdated).not.toContain('<a:srgbClr val="000000">');

    const switched = parse(source);
    expect(replaceShapeShadow(switched.xml, switched.shape, normalizeShapeShadow({
      kind: 'inner',
      color: { kind: 'scheme', value: 'accent2' },
      opacity: 0,
      blur: 0,
      angle: 0,
      distance: 0,
    }, 'x'), PART_URI)).toBe(true);
    const switchedUpdated = switched.xml.serialize();
    expect(switchedUpdated).toContain(INNER);
    expect(switchedUpdated).not.toContain('<a:outerShdw');

    const cleared = parse(source);
    expect(replaceShapeShadow(cleared.xml, cleared.shape, undefined, PART_URI)).toBe(true);
    const clearedUpdated = cleared.xml.serialize();
    expect(clearedUpdated).toContain('<a:effectLst xmlns:k="urn:keep"><a:blur/>');
    expect(clearedUpdated).not.toContain('outerShdw');
    for (const token of ['GLOW', 'PRESET', 'REFLECTION', 'SOFT', 'KEEP', 'rId9']) {
      expect(clearedUpdated).toContain(token);
    }

    const absent = parse(fixture(properties('<a:effectLst xmlns:k="urn:empty"/>')));
    expect(replaceShapeShadow(absent.xml, absent.shape, undefined, PART_URI)).toBe(false);
    expect(absent.xml.changed).toBe(false);
  });

  it('rejects every unsafe owner state before changing any byte', () => {
    const sources = [
      fixture(''),
      fixture('<p:spPr/><p:spPr/>'),
      fixture('<x:spPr xmlns:x="urn:wrong"/>'),
      fixture('<p:spPr><a:prstGeom prst="rect"/><a:effectDag/></p:spPr>'),
      fixture('<p:spPr><a:prstGeom prst="rect"/><a:effectLst/><a:effectLst/></p:spPr>'),
      fixture('<p:spPr><a:prstGeom prst="rect"/><x:effectLst xmlns:x="urn:wrong"/></p:spPr>'),
      fixture('<p:spPr><a:prstGeom prst="rect"/><a:effectLst>' +
        `${OUTER}${INNER}</a:effectLst></p:spPr>`),
      fixture('<p:spPr><a:prstGeom prst="rect"/><a:effectLst>' +
        '<a:reflection/><a:glow/></a:effectLst></p:spPr>'),
      fixture('<p:spPr><a:prstGeom prst="rect"/><a:effectLst>' +
        '<a:outerShdw blurRad="bad"><a:srgbClr val="000000"/>' +
        '</a:outerShdw></a:effectLst></p:spPr>'),
      fixture('<p:spPr><a:prstGeom prst="rect"/><a:effectLst>' +
        '<a:outerShdw sx="100000"><a:srgbClr val="000000"/>' +
        '</a:outerShdw></a:effectLst></p:spPr>'),
      fixture('<p:spPr><a:prstGeom prst="rect"/><a:effectLst>' +
        '<a:outerShdw><a:srgbClr val="000000"><a:alpha val="100001"/>' +
        '</a:srgbClr></a:outerShdw></a:effectLst></p:spPr>'),
      fixture('<p:spPr><a:prstGeom prst="rect"/><a:scene3d/><a:ln/></p:spPr>'),
      fixture('<p:spPr><a:prstGeom prst="rect"/><a:custGeom/></p:spPr>'),
    ];
    for (const source of sources) {
      const { xml, shape } = parse(source);
      expect(() => replaceShapeShadow(
        xml,
        shape,
        normalizeShapeShadow({ kind: 'outer' }, 'x'),
        PART_URI,
      ), source).toThrow(ModelParseError);
      expect(xml.changed, source).toBe(false);
      expect(xml.serialize(), source).toBe(source);
    }
  });
});
