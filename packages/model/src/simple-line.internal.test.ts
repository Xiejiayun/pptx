import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import { describe, expect, it } from 'vitest';
import {
  normalizeSimpleLine,
  readSimpleLine,
  renderSimpleLine,
  simpleLinesEqual,
  type SimpleLineDash,
} from './simple-line.internal.js';

const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const DASHES = [
  'solid',
  'dash',
  'dashDot',
  'lgDash',
  'lgDashDot',
  'lgDashDotDot',
  'sysDash',
  'sysDot',
] as const satisfies readonly SimpleLineDash[];

function parseLine(source: string): XmlElement {
  const xml = LosslessXmlDocument.parse(
    `<root xmlns:a="${DRAWING_NAMESPACE}" xmlns:d="${DRAWING_NAMESPACE}" ` +
    `xmlns:x="urn:wrong">${source}</root>`,
  );
  const root = xml.roots[0];
  const line = root?.children.find(
    (child): child is XmlElement => child.type === 'element',
  );
  if (!line) throw new Error('Fixture has no line');
  return line;
}

describe('simple line normalization', () => {
  it('normalizes defaults, quantized values, every dash, and detached colors', () => {
    expect(normalizeSimpleLine(undefined, 'Shape line')).toBeUndefined();
    expect(normalizeSimpleLine({ kind: 'none' }, 'Shape line')).toEqual({ kind: 'none' });

    const color = { kind: 'srgb', value: '#ff0000' };
    const input = { kind: 'line', color };
    const normalized = normalizeSimpleLine(input, 'Shape line');
    color.value = '000000';
    expect(normalized).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: 'FF0000' },
      width: 1,
      dash: 'solid',
    });
    expect(normalized).not.toBe(input);
    if (normalized?.kind === 'line') expect(normalized.color).not.toBe(color);

    expect(normalizeSimpleLine({
      kind: 'line',
      color: { kind: 'scheme', value: 'accent2' },
      transparency: 33.3334,
      width: 0.333333,
      dash: 'lgDashDotDot',
    }, 'Shape line')).toEqual({
      kind: 'line',
      color: { kind: 'scheme', value: 'accent2' },
      transparency: 33.333,
      width: 4_233 / 12_700,
      dash: 'lgDashDotDot',
    });

    expect(normalizeSimpleLine({
      kind: 'line',
      color: { kind: 'srgb', value: '112233' },
      transparency: 0,
      width: 0,
      dash: 'sysDot',
    }, 'Shape line')).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: '112233' },
      transparency: 0,
      width: 0,
      dash: 'sysDot',
    });

    expect(normalizeSimpleLine({
      kind: 'line',
      color: { kind: 'srgb', value: 'FFFFFF' },
      transparency: 100,
      width: 1_584,
    }, 'Shape line')).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: 'FFFFFF' },
      transparency: 100,
      width: 1_584,
      dash: 'solid',
    });

    for (const dash of DASHES) {
      expect(normalizeSimpleLine({
        kind: 'line',
        color: { kind: 'srgb', value: 'ABCDEF' },
        dash,
      }, 'Shape line')).toMatchObject({ dash });
    }

    const nullColor = Object.create(null) as Record<string, unknown>;
    nullColor.kind = 'scheme';
    nullColor.value = 'accent4';
    const nullLine = Object.create(null) as Record<string, unknown>;
    nullLine.kind = 'line';
    nullLine.color = nullColor;
    expect(normalizeSimpleLine(nullLine, 'Shape line')).toEqual({
      kind: 'line',
      color: { kind: 'scheme', value: 'accent4' },
      width: 1,
      dash: 'solid',
    });
  });

  it('rejects non-ordinary, inherited, unknown, and alias-bearing values', () => {
    class Line {
      kind = 'none';
    }
    for (const value of [null, false, 1, 'line', [], new Date(), new Line()]) {
      expect(() => normalizeSimpleLine(value, 'Shape line')).toThrow(TypeError);
    }
    expect(() => normalizeSimpleLine(Object.create({ kind: 'none' }), 'Shape line'))
      .toThrow(TypeError);
    expect(() => normalizeSimpleLine({ kind: 'none', width: undefined }, 'Shape line'))
      .toThrow(/unsupported property width/);

    for (const alias of ['type', 'alpha', 'dashType', 'lineDash', 'lineHead', 'lineTail']) {
      expect(() => normalizeSimpleLine({
        kind: 'line',
        color: { kind: 'srgb', value: 'FFFFFF' },
        [alias]: undefined,
      }, 'Shape line'), alias).toThrow(new RegExp(`unsupported property ${alias}`));
    }
    expect(() => normalizeSimpleLine({ kind: 'none', [Symbol('unsafe')]: true }, 'Shape line'))
      .toThrow(/unsupported property/);
  });

  it('rejects accessors without invoking them', () => {
    let calls = 0;
    const kindGetter = Object.defineProperty({}, 'kind', {
      enumerable: true,
      get() {
        calls += 1;
        return 'none';
      },
    });
    const colorGetter = Object.defineProperty({ kind: 'line' }, 'color', {
      enumerable: true,
      get() {
        calls += 1;
        return { kind: 'srgb', value: 'FFFFFF' };
      },
    });
    const widthGetter = Object.defineProperty({
      kind: 'line',
      color: { kind: 'srgb', value: 'FFFFFF' },
    }, 'width', {
      enumerable: true,
      get() {
        calls += 1;
        return 1;
      },
    });
    for (const value of [kindGetter, colorGetter, widthGetter]) {
      expect(() => normalizeSimpleLine(value, 'Shape line')).toThrow(/data property/);
    }
    expect(calls).toBe(0);
  });

  it('rejects invalid kind, color, transparency, width, and dash values', () => {
    for (const value of [
      {},
      { kind: 'solid' },
      { kind: 'line' },
      { kind: 'line', color: null },
      { kind: 'line', color: { kind: 'rgb', value: 'FFFFFF' } },
      { kind: 'line', color: { kind: 'srgb', value: 'FFF' } },
      { kind: 'line', color: { kind: 'scheme', value: 'unknown' } },
    ]) {
      expect(() => normalizeSimpleLine(value, 'Shape line')).toThrow(TypeError);
    }

    for (const transparency of [
      '50',
      Number.NaN,
      Number.POSITIVE_INFINITY,
      Number.NEGATIVE_INFINITY,
    ]) {
      expect(() => normalizeSimpleLine({
        kind: 'line',
        color: { kind: 'srgb', value: 'FFFFFF' },
        transparency,
      }, 'Shape line')).toThrow(TypeError);
    }
    for (const transparency of [-0.001, 100.001]) {
      expect(() => normalizeSimpleLine({
        kind: 'line',
        color: { kind: 'srgb', value: 'FFFFFF' },
        transparency,
      }, 'Shape line')).toThrow(RangeError);
    }

    for (const width of [
      '1',
      Number.NaN,
      Number.POSITIVE_INFINITY,
      Number.NEGATIVE_INFINITY,
    ]) {
      expect(() => normalizeSimpleLine({
        kind: 'line',
        color: { kind: 'srgb', value: 'FFFFFF' },
        width,
      }, 'Shape line')).toThrow(TypeError);
    }
    for (const width of [-0.001, 1_584.001]) {
      expect(() => normalizeSimpleLine({
        kind: 'line',
        color: { kind: 'srgb', value: 'FFFFFF' },
        width,
      }, 'Shape line')).toThrow(RangeError);
    }

    for (const dash of ['', 'dot', 'DASH', 1, null]) {
      expect(() => normalizeSimpleLine({
        kind: 'line',
        color: { kind: 'srgb', value: 'FFFFFF' },
        dash,
      }, 'Shape line')).toThrow(TypeError);
    }
  });
});

describe('simple line codec', () => {
  it('renders deterministic none, color, alpha, and every preset dash child', () => {
    expect(renderSimpleLine({ kind: 'none' }, 'a:')).toBe('<a:noFill/>');
    expect(renderSimpleLine({
      kind: 'line',
      color: { kind: 'srgb', value: 'FF0000' },
      transparency: 50,
      width: 2.5,
      dash: 'dashDot',
    }, 'a:')).toBe(
      '<a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000"/>' +
      '</a:srgbClr></a:solidFill><a:prstDash val="dashDot"/>',
    );

    for (const dash of DASHES) {
      expect(renderSimpleLine({
        kind: 'line',
        color: { kind: 'scheme', value: 'accent3' },
        width: 1,
        dash,
      }, 'd:')).toBe(
        '<d:solidFill><d:schemeClr val="accent3"/></d:solidFill>' +
        `<d:prstDash val="${dash}"/>`,
      );
    }
  });

  it('reads none and normalized solid lines with defaults and alternate prefixes', () => {
    expect(readSimpleLine(parseLine('<a:ln><a:noFill/></a:ln>'), 'a:'))
      .toEqual({ kind: 'none' });
    expect(readSimpleLine(parseLine(
      '<a:ln cap="flat"><a:noFill/><a:round/><a:headEnd type="triangle"/>' +
      '<a:tailEnd type="arrow"/><a:extLst><a:ext uri="urn:keep"/></a:extLst></a:ln>',
    ), 'a:')).toEqual({ kind: 'none' });
    expect(readSimpleLine(parseLine(
      '<a:ln><a:solidFill><a:srgbClr val="ff0000"/></a:solidFill></a:ln>',
    ), 'a:')).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: 'FF0000' },
      width: 1,
      dash: 'solid',
    });
    expect(readSimpleLine(parseLine(
      '<d:ln w="0"><d:solidFill><d:schemeClr val="accent2">' +
      '<d:alpha val="75000"/></d:schemeClr></d:solidFill>' +
      '<d:prstDash val="lgDashDot"/></d:ln>',
    ), 'd:')).toEqual({
      kind: 'line',
      color: { kind: 'scheme', value: 'accent2' },
      transparency: 25,
      width: 0,
      dash: 'lgDashDot',
    });
    expect(readSimpleLine(parseLine(
      '<a:ln w="31750" cap="flat" cmpd="sng" algn="ctr">' +
      '<a:solidFill><a:srgbClr val="112233"/></a:solidFill>' +
      '<a:prstDash val="sysDot"/><a:round/><a:headEnd type="triangle"/>' +
      '<a:tailEnd type="arrow"/><a:extLst><a:ext uri="urn:keep"/></a:extLst></a:ln>',
    ), 'a:')).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: '112233' },
      width: 2.5,
      dash: 'sysDot',
    });

    for (const dash of DASHES) {
      expect(readSimpleLine(parseLine(
        `<a:ln w="20116800"><a:solidFill><a:srgbClr val="ABCDEF"/>` +
        `</a:solidFill><a:prstDash val="${dash}"/></a:ln>`,
      ), 'a:'), dash).toEqual({
        kind: 'line',
        color: { kind: 'srgb', value: 'ABCDEF' },
        width: 1_584,
        dash,
      });
    }
  });

  it('returns detached snapshots without mutating source XML', () => {
    const source = '<a:ln w="12700"><a:solidFill><a:srgbClr val="ABCDEF"/>' +
      '</a:solidFill><a:prstDash val="dash"/></a:ln>';
    const line = parseLine(source);
    const first = readSimpleLine(line, 'a:');
    const second = readSimpleLine(line, 'a:');
    expect(first).toEqual(second);
    expect(first).not.toBe(second);
    if (first?.kind === 'line' && second?.kind === 'line') {
      expect(first.color).not.toBe(second.color);
    }
  });

  it('returns undefined for empty, unsupported, malformed, conflicting, or lookalike state', () => {
    const fixtures = [
      '<a:ln/>',
      '<a:ln><a:noFill/><a:prstDash val="solid"/></a:ln>',
      '<a:ln w="12700"><a:noFill/></a:ln>',
      '<a:ln><a:gradFill/></a:ln>',
      '<a:ln><a:blipFill/></a:ln>',
      '<a:ln><a:pattFill/></a:ln>',
      '<a:ln><a:grpFill/></a:ln>',
      '<a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>' +
        '<a:custDash><a:ds d="1" sp="1"/></a:custDash></a:ln>',
      '<a:ln><a:noFill/><a:solidFill><a:srgbClr val="FF0000"/>' +
        '</a:solidFill></a:ln>',
      '<a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>' +
        '<a:prstDash val="dash"/><a:prstDash val="solid"/></a:ln>',
      '<a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>' +
        '<a:prstDash/></a:ln>',
      '<a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>' +
        '<a:prstDash val="dot"/></a:ln>',
      '<a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>' +
        '<a:prstDash val="dash" custom="x"/></a:ln>',
      '<a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>' +
        '<a:prstDash val="dash"><a:ext/></a:prstDash></a:ln>',
      '<a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>' +
        '<x:prstDash val="dash"/></a:ln>',
      '<a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>' +
        '<a:prstDash xmlns:a="urn:wrong" val="dash"/></a:ln>',
      '<a:ln><a:solidFill/></a:ln>',
      '<a:ln w=""><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>',
      '<a:ln w="1.5"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>',
      '<a:ln w="-1"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>',
      '<a:ln w="20116801"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>',
      '<a:ln custom="x"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>',
      '<a:ln cap="round"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>',
      '<a:ln cmpd="dbl"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>',
      '<a:ln algn="in"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>',
      '<a:ln><x:solidFill><x:srgbClr val="FF0000"/></x:solidFill></a:ln>',
      '<a:ln><a:solidFill xmlns:a="urn:wrong"><a:srgbClr val="FF0000"/>' +
        '</a:solidFill></a:ln>',
      '<a:ln xmlns:a="urn:wrong"><a:solidFill><a:srgbClr val="FF0000"/>' +
        '</a:solidFill></a:ln>',
      '<x:ln><x:solidFill><x:srgbClr val="FF0000"/></x:solidFill></x:ln>',
    ];
    for (const source of fixtures) {
      expect(readSimpleLine(parseLine(source), 'a:'), source).toBeUndefined();
    }
  });

  it('compares kind, color, transparency presence, width, and dash exactly', () => {
    const red = {
      kind: 'line',
      color: { kind: 'srgb', value: 'FF0000' },
      width: 1,
      dash: 'solid',
    } as const;
    expect(simpleLinesEqual({ kind: 'none' }, { kind: 'none' })).toBe(true);
    expect(simpleLinesEqual(undefined, { kind: 'none' })).toBe(false);
    expect(simpleLinesEqual(red, { ...red })).toBe(true);
    expect(simpleLinesEqual(red, { ...red, transparency: 0 })).toBe(false);
    expect(simpleLinesEqual(red, { ...red, width: 2 })).toBe(false);
    expect(simpleLinesEqual(red, { ...red, dash: 'dash' })).toBe(false);
    expect(simpleLinesEqual(red, {
      ...red,
      color: { kind: 'scheme', value: 'accent1' },
    })).toBe(false);
  });
});
