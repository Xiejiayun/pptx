import { LosslessXmlDocument } from '@pptx/lossless-xml';
import { describe, expect, it } from 'vitest';
import { ModelParseError } from './errors.js';
import {
  normalizeShapeArrows,
  readShapeArrows,
  renderShapeArrows,
  replaceShapeArrows,
  shapeArrowsEqual,
} from './shape-arrows.internal.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const PART_URI = '/ppt/slides/slide1.xml';

function fixture(
  properties: string,
  options: {
    readonly rootName?: string;
    readonly presentationNamespace?: string;
  } = {},
): string {
  const rootName = options.rootName ?? 'p:sp';
  return `<${rootName} xmlns:p="${options.presentationNamespace ?? PRESENTATION_NAMESPACE}" ` +
    `xmlns:a="${DRAWING_NAMESPACE}">` +
    '<p:nvSpPr><p:cNvPr id="7" name="Keep"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
    `${properties}<p:txBody><a:bodyPr/><a:p><a:r><a:t>KEEP</a:t></a:r></a:p></p:txBody>` +
    `</${rootName}>`;
}

function properties(
  line = '<a:ln/>',
  fill = '<a:noFill/>',
  geometry = '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>',
): string {
  return '<p:spPr keep="PROPERTIES"><a:xfrm><a:off x="1" y="2"/>' +
    `<a:ext cx="3" cy="4"/></a:xfrm>${geometry}${fill}${line}` +
    '<a:effectLst/><a:scene3d/><a:sp3d/>' +
    '<a:extLst><a:ext uri="urn:keep"><x:opaque xmlns:x="urn:test"/></a:ext></a:extLst>' +
    '</p:spPr>';
}

function parse(source: string) {
  const xml = LosslessXmlDocument.parse(source);
  const shape = xml.roots[0];
  if (!shape) throw new Error('Fixture has no shape');
  return { xml, shape };
}

describe('shape arrow normalization', () => {
  it('normalizes, freezes, and detaches empty, partial, and complete snapshots', () => {
    expect(normalizeShapeArrows(undefined, 'Shape arrows')).toBeUndefined();
    const empty = normalizeShapeArrows({}, 'Shape arrows');
    expect(empty).toEqual({});
    expect(Object.isFrozen(empty)).toBe(true);
    expect(normalizeShapeArrows({ begin: undefined, end: 'arrow' }, 'Shape arrows'))
      .toEqual({ end: 'arrow' });

    const input: { begin: string; end: string | undefined } = {
      begin: 'triangle',
      end: 'oval',
    };
    const normalized = normalizeShapeArrows(input, 'Shape arrows');
    input.begin = 'diamond';
    input.end = undefined;
    expect(normalized).toEqual({ begin: 'triangle', end: 'oval' });
    expect(Object.isFrozen(normalized)).toBe(true);

    const nullPrototype = Object.create(null) as Record<string, unknown>;
    nullPrototype.begin = 'none';
    nullPrototype.end = 'stealth';
    expect(normalizeShapeArrows(nullPrototype, 'Shape arrows'))
      .toEqual({ begin: 'none', end: 'stealth' });
  });

  it('accepts exactly the six public tokens on either side', () => {
    const types = ['none', 'arrow', 'diamond', 'oval', 'stealth', 'triangle'] as const;
    for (const type of types) {
      expect(normalizeShapeArrows({ begin: type }, 'Shape arrows')).toEqual({ begin: type });
      expect(normalizeShapeArrows({ end: type }, 'Shape arrows')).toEqual({ end: type });
    }
  });

  it('rejects unsafe objects, aliases, accessors, symbols, and invalid tokens', () => {
    class ArrowOptions {
      begin = 'arrow';
    }
    const inherited = Object.create({ begin: 'arrow' }) as Record<string, unknown>;
    const symbol = { begin: 'arrow', [Symbol('arrow')]: 'triangle' };
    const invalidValues = [
      null,
      true,
      1,
      'arrow',
      [],
      new Date(),
      new ArrowOptions(),
      inherited,
      symbol,
      { unknown: 'arrow' },
      { beginArrowType: 'arrow' },
      { endArrowType: 'arrow' },
      { lineHead: 'arrow' },
      { lineTail: 'arrow' },
      { begin: null },
      { begin: '' },
      { begin: 'Arrow' },
      { begin: 'bogus' },
      { begin: 1 },
      { end: '' },
      { end: 'TRIANGLE' },
    ];
    for (const value of invalidValues) {
      expect(() => normalizeShapeArrows(value, 'Shape arrows'), JSON.stringify(value))
        .toThrow(TypeError);
    }

    let getterInvoked = false;
    const accessor = Object.defineProperty({}, 'begin', {
      enumerable: true,
      get() {
        getterInvoked = true;
        return 'arrow';
      },
    });
    expect(() => normalizeShapeArrows(accessor, 'Shape arrows')).toThrow(TypeError);
    expect(getterInvoked).toBe(false);
  });

  it('renders canonical head-before-tail XML and compares only endpoint types', () => {
    expect(renderShapeArrows(undefined, 'a:')).toBe('');
    expect(renderShapeArrows({}, 'a:')).toBe('');
    expect(renderShapeArrows({ begin: 'triangle' }, 'a:'))
      .toBe('<a:headEnd type="triangle"/>');
    expect(renderShapeArrows({ end: 'arrow' }, 'd:'))
      .toBe('<d:tailEnd type="arrow"/>');
    expect(renderShapeArrows({ begin: 'triangle', end: 'arrow' }, 'a:'))
      .toBe('<a:headEnd type="triangle"/><a:tailEnd type="arrow"/>');
    expect(shapeArrowsEqual(undefined, undefined)).toBe(true);
    expect(shapeArrowsEqual(undefined, {})).toBe(true);
    expect(shapeArrowsEqual({ begin: 'none' }, {})).toBe(false);
    expect(shapeArrowsEqual(
      { begin: 'triangle', end: 'arrow' },
      { begin: 'triangle', end: 'arrow' },
    )).toBe(true);
  });
});

describe('shape arrow reader', () => {
  it('reads detached begin, end, explicit none, legal sizes, and alternate prefixes', () => {
    const cases = [
      {
        source: fixture(properties(
          '<a:ln><a:headEnd type="triangle"/><a:tailEnd type="arrow"/></a:ln>',
        )),
        expected: { begin: 'triangle', end: 'arrow' },
      },
      {
        source: fixture(properties(
          '<a:ln><a:headEnd type="none" w="med" len="med"/></a:ln>',
        )),
        expected: { begin: 'none' },
      },
      {
        source: fixture(properties(
          '<a:ln><a:tailEnd type="stealth" w="sm" len="lg"/></a:ln>',
        )),
        expected: { end: 'stealth' },
      },
      {
        source:
          `<q:sp xmlns:q="${PRESENTATION_NAMESPACE}" xmlns:d="${DRAWING_NAMESPACE}">` +
          '<q:spPr><d:prstGeom prst="line"><d:avLst/></d:prstGeom><d:noFill/>' +
          '<d:ln><d:headEnd type="diamond" w="lg" len="sm"/>' +
          '<d:tailEnd type="oval"/></d:ln></q:spPr></q:sp>',
        expected: { begin: 'diamond', end: 'oval' },
      },
    ];

    for (const { source, expected } of cases) {
      const { xml, shape } = parse(source);
      const first = readShapeArrows(xml, shape);
      const second = readShapeArrows(xml, shape);
      expect(first, source).toEqual(expected);
      expect(second, source).toEqual(expected);
      expect(first).not.toBe(second);
      expect(Object.isFrozen(first)).toBe(true);
      expect(xml.changed).toBe(false);
      expect(xml.serialize()).toBe(source);
    }
  });

  it('returns undefined for absent, empty, nested, malformed, ambiguous, or unsafe state', () => {
    const lines = [
      '',
      '<a:ln/>',
      '<a:ln><a:round/></a:ln>',
      '<a:ln><a:extLst><a:headEnd type="arrow"/></a:extLst></a:ln>',
      '<a:ln><a:headEnd/></a:ln>',
      '<a:ln><a:headEnd type="bogus"/></a:ln>',
      '<a:ln><a:headEnd a:type="arrow"/></a:ln>',
      '<a:ln><a:headEnd type="arrow" type="triangle"/></a:ln>',
      '<a:ln><a:headEnd type="arrow" w="xl"/></a:ln>',
      '<a:ln><a:headEnd type="arrow" w="sm" w="lg"/></a:ln>',
      '<a:ln><a:headEnd type="arrow" data-extra="x"/></a:ln>',
      '<a:ln><a:headEnd type="arrow"><a:ext/></a:headEnd></a:ln>',
      '<a:ln><a:headEnd type="arrow">text</a:headEnd></a:ln>',
      '<a:ln><x:headEnd xmlns:x="urn:wrong" type="arrow"/></a:ln>',
      '<a:ln><a:headEnd type="arrow"/><a:headEnd type="triangle"/></a:ln>',
      '<a:ln><a:tailEnd type="arrow"/><a:tailEnd type="triangle"/></a:ln>',
      '<a:ln><a:tailEnd type="arrow"/><a:headEnd type="triangle"/></a:ln>',
      '<a:ln><a:extLst/><a:headEnd type="arrow"/></a:ln>',
      '<a:ln><a:headEnd type="arrow"/><a:round/></a:ln>',
      '<a:ln><x:unknown xmlns:x="urn:test"/><a:headEnd type="arrow"/></a:ln>',
    ];
    for (const line of lines) {
      const source = fixture(properties(line));
      const { xml, shape } = parse(source);
      expect(readShapeArrows(xml, shape), line).toBeUndefined();
      expect(xml.changed, line).toBe(false);
      expect(xml.serialize(), line).toBe(source);
    }

    for (const source of [
      fixture('<p:spPr/><p:spPr/>'),
      fixture(properties('<a:ln/><a:ln/>')),
      fixture(properties('<x:ln xmlns:x="urn:wrong"><x:headEnd type="arrow"/></x:ln>')),
      fixture(properties('<a:ln><a:headEnd type="arrow"/></a:ln>'), { rootName: 'p:pic' }),
      fixture(properties('<a:ln><a:headEnd type="arrow"/></a:ln>'), {
        presentationNamespace: 'urn:wrong',
      }),
    ]) {
      const { xml, shape } = parse(source);
      expect(readShapeArrows(xml, shape)).toBeUndefined();
      expect(xml.serialize()).toBe(source);
    }
  });
});

describe('shape arrow replacement', () => {
  it('preserves exact bytes for same-value assignments including legal sizes', () => {
    const source = fixture(properties(
      '<a:ln w="31750"><a:solidFill><a:srgbClr val="112233"/></a:solidFill>' +
      '<a:prstDash val="dashDot"/><a:round/>' +
      '<a:headEnd type="triangle" w="lg" len="sm"/>' +
      '<a:tailEnd type="arrow" w="med" len="med"/><a:extLst/></a:ln>',
    ));
    const { xml, shape } = parse(source);
    expect(replaceShapeArrows(xml, shape, {
      begin: 'triangle',
      end: 'arrow',
    }, PART_URI)).toBe(false);
    expect(xml.changed).toBe(false);
    expect(xml.serialize()).toBe(source);
  });

  it('replaces only type values and preserves advanced line state and legal sizes', () => {
    const source = fixture(properties(
      '<a:ln w="31750" cap="flat" cmpd="sng" algn="ctr" data-keep="LINE">' +
      '<a:gradFill><a:gsLst/></a:gradFill>' +
      '<a:custDash><a:ds d="1" sp="1"/></a:custDash><a:round/>' +
      '<a:headEnd type="triangle" w="lg" len="sm"/>' +
      '<a:tailEnd type="arrow" w="med" len="med"/>' +
      '<a:extLst><a:ext uri="urn:line"><x:keep xmlns:x="urn:line"/>' +
      '</a:ext></a:extLst></a:ln>',
    ));
    const { xml, shape } = parse(source);
    expect(replaceShapeArrows(xml, shape, {
      begin: 'diamond',
      end: 'oval',
    }, PART_URI)).toBe(true);
    const updated = xml.serialize();
    expect(updated).toContain(
      '<a:ln w="31750" cap="flat" cmpd="sng" algn="ctr" data-keep="LINE">' +
      '<a:gradFill><a:gsLst/></a:gradFill>' +
      '<a:custDash><a:ds d="1" sp="1"/></a:custDash><a:round/>' +
      '<a:headEnd type="diamond" w="lg" len="sm"/>' +
      '<a:tailEnd type="oval" w="med" len="med"/>' +
      '<a:extLst><a:ext uri="urn:line"><x:keep xmlns:x="urn:line"/>',
    );
    expect(updated).toContain('<p:cNvPr id="7" name="Keep"/>');
    expect(updated).toContain('<a:t>KEEP</a:t>');
  });

  it('applies whole-replacement side semantics without rewriting line siblings', () => {
    const base = '<a:ln w="12700"><a:solidFill><a:srgbClr val="112233"/>' +
      '</a:solidFill><a:prstDash val="solid"/><a:bevel/>' +
      '<a:headEnd type="triangle" w="lg"/><a:tailEnd type="arrow" len="sm"/>' +
      '<a:extLst><a:ext uri="urn:keep"/></a:extLst></a:ln>';
    const cases = [
      {
        target: { begin: 'diamond' },
        expected: '<a:bevel/><a:headEnd type="diamond" w="lg"/><a:extLst>',
        absent: '<a:tailEnd',
      },
      {
        target: { end: 'oval' },
        expected: '<a:bevel/><a:tailEnd type="oval" len="sm"/><a:extLst>',
        absent: '<a:headEnd',
      },
      {
        target: {},
        expected: '<a:prstDash val="solid"/><a:bevel/><a:extLst>',
        absent: '<a:headEnd',
      },
    ] as const;
    for (const { target, expected, absent } of cases) {
      const source = fixture(properties(base));
      const { xml, shape } = parse(source);
      expect(replaceShapeArrows(xml, shape, target, PART_URI)).toBe(true);
      const updated = xml.serialize();
      expect(updated, JSON.stringify(target)).toContain(expected);
      expect(updated, JSON.stringify(target)).not.toContain(absent);
      if (Object.keys(target).length === 0) {
        expect(updated).not.toContain('<a:tailEnd');
      }
      expect(updated).toContain('<a:ln w="12700">');
      expect(updated).toContain('<a:solidFill><a:srgbClr val="112233"/></a:solidFill>');
      expect(updated).toContain('<a:ext uri="urn:keep"/>');
    }
  });

  it('adds endpoints to self-closing and existing lines in schema order', () => {
    const cases = [
      {
        line: '<a:ln/>',
        target: { begin: 'triangle', end: 'arrow' },
        expected: '<a:ln><a:headEnd type="triangle"/><a:tailEnd type="arrow"/></a:ln>',
      },
      {
        line: '<a:ln><a:tailEnd type="arrow"/><a:extLst/></a:ln>',
        target: { begin: 'diamond', end: 'arrow' },
        expected: '<a:headEnd type="diamond"/><a:tailEnd type="arrow"/><a:extLst/>',
      },
      {
        line: '<a:ln><a:round/><a:headEnd type="triangle"/><a:extLst/></a:ln>',
        target: { begin: 'triangle', end: 'oval' },
        expected: '<a:round/><a:headEnd type="triangle"/><a:tailEnd type="oval"/><a:extLst/>',
      },
      {
        line: '<a:ln><a:round/><a:extLst/></a:ln>',
        target: { begin: 'stealth', end: 'none' },
        expected: '<a:round/><a:headEnd type="stealth"/><a:tailEnd type="none"/><a:extLst/>',
      },
    ] as const;
    for (const { line, target, expected } of cases) {
      const source = fixture(properties(line));
      const { xml, shape } = parse(source);
      expect(replaceShapeArrows(xml, shape, target, PART_URI)).toBe(true);
      expect(xml.serialize(), line).toContain(expected);
    }
  });

  it('swaps a single existing side without overlapping local edits', () => {
    const cases = [
      {
        line: '<a:ln><a:headEnd type="triangle" w="lg"/><a:extLst/></a:ln>',
        target: { end: 'arrow' },
        expected: '<a:ln><a:tailEnd type="arrow"/><a:extLst/></a:ln>',
      },
      {
        line: '<a:ln><a:tailEnd type="arrow" len="sm"/><a:extLst/></a:ln>',
        target: { begin: 'diamond' },
        expected: '<a:ln><a:headEnd type="diamond"/><a:extLst/></a:ln>',
      },
    ] as const;
    for (const { line, target, expected } of cases) {
      const source = fixture(properties(line));
      const { xml, shape } = parse(source);
      expect(replaceShapeArrows(xml, shape, target, PART_URI)).toBe(true);
      expect(xml.serialize(), line).toContain(expected);
    }
  });

  it('inserts a missing line after shape fill and before effects with safe prefixes', () => {
    const cases = [
      {
        source: fixture(properties('', '<a:noFill/>')),
        expected:
          '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>' +
          '<a:ln><a:headEnd type="triangle"/><a:tailEnd type="arrow"/></a:ln>' +
          '<a:effectLst/>',
      },
      {
        source: fixture(properties('', '')),
        expected:
          '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
          '<a:ln><a:headEnd type="triangle"/><a:tailEnd type="arrow"/></a:ln>' +
          '<a:effectLst/>',
      },
      {
        source:
          `<q:sp xmlns:q="${PRESENTATION_NAMESPACE}" xmlns:d="${DRAWING_NAMESPACE}">` +
          '<q:spPr><d:prstGeom prst="rect"><d:avLst/></d:prstGeom><d:noFill/>' +
          '<d:effectLst/></q:spPr></q:sp>',
        expected:
          '<d:noFill/><d:ln><d:headEnd type="triangle"/>' +
          '<d:tailEnd type="arrow"/></d:ln><d:effectLst/>',
      },
      {
        source:
          `<p:sp xmlns:p="${PRESENTATION_NAMESPACE}"><p:spPr>` +
          `<d:prstGeom xmlns:d="${DRAWING_NAMESPACE}" prst="rect"><d:avLst/>` +
          '</d:prstGeom><p:extLst/></p:spPr></p:sp>',
        expected:
          `</d:prstGeom><d:ln xmlns:d="${DRAWING_NAMESPACE}">` +
          '<d:headEnd type="triangle"/><d:tailEnd type="arrow"/></d:ln><p:extLst/>',
      },
    ];
    for (const { source, expected } of cases) {
      const { xml, shape } = parse(source);
      expect(replaceShapeArrows(xml, shape, {
        begin: 'triangle',
        end: 'arrow',
      }, PART_URI), source).toBe(true);
      expect(xml.serialize(), source).toContain(expected);
    }
  });

  it('rejects malformed, repeated, wrong-namespace, and unsafe state before patching', () => {
    const sources = [
      fixture(''),
      fixture('<p:spPr/><p:spPr/>'),
      fixture('<x:spPr xmlns:x="urn:wrong"><a:ln/></x:spPr>'),
      fixture(properties('<a:ln/><a:ln/>')),
      fixture(properties('<x:ln xmlns:x="urn:wrong"/>')),
      fixture(properties('<a:ln><a:headEnd/></a:ln>')),
      fixture(properties('<a:ln><a:headEnd type="bogus"/></a:ln>')),
      fixture(properties('<a:ln><a:headEnd type="arrow" w="xl"/></a:ln>')),
      fixture(properties('<a:ln><a:headEnd type="arrow"><a:ext/></a:headEnd></a:ln>')),
      fixture(properties(
        '<a:ln><a:extLst><a:headEnd type="arrow"/></a:extLst></a:ln>',
      )),
      fixture(properties('<a:ln><x:headEnd xmlns:x="urn:wrong" type="arrow"/></a:ln>')),
      fixture(properties(
        '<a:ln><a:headEnd type="arrow"/><a:headEnd type="triangle"/></a:ln>',
      )),
      fixture(properties(
        '<a:ln><a:tailEnd type="arrow"/><a:headEnd type="triangle"/></a:ln>',
      )),
      fixture(properties('<a:ln><a:extLst/><a:headEnd type="arrow"/></a:ln>')),
      fixture(properties('<a:ln><a:headEnd type="arrow"/><a:round/></a:ln>')),
      fixture(properties('<a:ln><x:unknown xmlns:x="urn:test"/></a:ln>')),
      fixture(properties('', '<a:noFill/><a:solidFill><a:srgbClr val="FF0000"/>' +
        '</a:solidFill>')),
      fixture(properties('', '', '')),
      fixture(properties('', '', '<a:prstGeom prst="rect"/><a:custGeom/>')),
      fixture('<p:spPr><a:effectLst/><a:prstGeom prst="rect"/>' +
        '<a:noFill/></p:spPr>'),
    ];
    for (const source of sources) {
      const { xml, shape } = parse(source);
      expect(() => replaceShapeArrows(xml, shape, {
        begin: 'triangle',
      }, PART_URI), source).toThrow(ModelParseError);
      expect(xml.changed, source).toBe(false);
      expect(xml.serialize(), source).toBe(source);
    }
  });

  it('treats clear on absent or already-empty endpoint state as an exact no-op', () => {
    for (const source of [
      fixture('<p:spPr keep="EMPTY"/>'),
      fixture(properties('<a:ln><a:round/><a:extLst/></a:ln>')),
    ]) {
      const { xml, shape } = parse(source);
      expect(replaceShapeArrows(xml, shape, undefined, PART_URI)).toBe(false);
      expect(xml.changed).toBe(false);
      expect(xml.serialize()).toBe(source);
    }
  });
});
