import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import { describe, expect, it } from 'vitest';
import {
  normalizeSimpleFill,
  readSimpleFillChoice,
  renderSimpleFill,
  simpleFillsEqual,
} from './simple-fill.internal.js';

function parseChoice(source: string): XmlElement {
  const xml = LosslessXmlDocument.parse(
    `<root xmlns:a="urn:drawing" xmlns:d="urn:drawing">${source}</root>`,
  );
  const root = xml.roots[0];
  const choice = root?.children.find(
    (child): child is XmlElement => child.type === 'element',
  );
  if (!choice) throw new Error('Fixture has no fill choice');
  return choice;
}

describe('simple fill normalization', () => {
  it('normalizes none, sRGB, scheme, and transparency into detached values', () => {
    expect(normalizeSimpleFill(undefined, 'Fill')).toBeUndefined();
    expect(normalizeSimpleFill({ kind: 'none' }, 'Fill')).toEqual({ kind: 'none' });

    const color = { kind: 'srgb', value: '#ff0000' };
    const input = { kind: 'solid', color, transparency: 33.3334 };
    const normalized = normalizeSimpleFill(input, 'Fill');
    color.value = '000000';
    input.transparency = 1;
    expect(normalized).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: 'FF0000' },
      transparency: 33.333,
    });
    expect(normalized).not.toBe(input);
    if (normalized?.kind === 'solid') expect(normalized.color).not.toBe(color);

    const nullPrototypeColor = Object.create(null) as Record<string, unknown>;
    nullPrototypeColor.kind = 'scheme';
    nullPrototypeColor.value = 'accent2';
    const nullPrototypeFill = Object.create(null) as Record<string, unknown>;
    nullPrototypeFill.kind = 'solid';
    nullPrototypeFill.color = nullPrototypeColor;
    nullPrototypeFill.transparency = 0;
    expect(normalizeSimpleFill(nullPrototypeFill, 'Fill')).toEqual({
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent2' },
      transparency: 0,
    });
  });

  it('rejects non-ordinary containers, inherited values, and unknown properties', () => {
    class Fill {
      kind = 'none';
    }
    for (const value of [null, false, 1, 'fill', [], new Date(), new Fill()]) {
      expect(() => normalizeSimpleFill(value, 'Fill')).toThrow(TypeError);
    }
    expect(() => normalizeSimpleFill(Object.create({ kind: 'none' }), 'Fill'))
      .toThrow(TypeError);
    expect(() => normalizeSimpleFill({ kind: 'none', color: undefined }, 'Fill'))
      .toThrow(/unsupported property color/);
    expect(() => normalizeSimpleFill({ kind: 'solid', color: { kind: 'srgb', value: 'FFFFFF' }, alpha: 10 }, 'Fill'))
      .toThrow(/unsupported property alpha/);
    expect(() => normalizeSimpleFill({ kind: 'none', [Symbol('unsafe')]: true }, 'Fill'))
      .toThrow(/unsupported property/);
  });

  it('rejects accessors without invoking them', () => {
    let calls = 0;
    const fillGetter = Object.defineProperty({}, 'kind', {
      enumerable: true,
      get() {
        calls += 1;
        return 'none';
      },
    });
    const fillSetter = Object.defineProperty({}, 'kind', {
      enumerable: true,
      set(_value: unknown) {
        calls += 1;
      },
    });
    const colorGetter = Object.defineProperty({ kind: 'solid' }, 'color', {
      enumerable: true,
      get() {
        calls += 1;
        return { kind: 'srgb', value: 'FFFFFF' };
      },
    });
    const colorValueGetter = Object.defineProperty({ kind: 'srgb' }, 'value', {
      enumerable: true,
      get() {
        calls += 1;
        return 'FFFFFF';
      },
    });
    for (const value of [
      fillGetter,
      fillSetter,
      colorGetter,
      { kind: 'solid', color: colorValueGetter },
    ]) {
      expect(() => normalizeSimpleFill(value, 'Fill')).toThrow(/data property/);
    }
    expect(calls).toBe(0);
  });

  it('rejects invalid kinds, colors, and transparency', () => {
    for (const value of [
      {},
      { kind: 'gradient' },
      { kind: 'solid' },
      { kind: 'solid', color: null },
      { kind: 'solid', color: { kind: 'rgb', value: 'FFFFFF' } },
      { kind: 'solid', color: { kind: 'srgb', value: 'FFF' } },
      { kind: 'solid', color: { kind: 'srgb', value: 'GGGGGG' } },
      { kind: 'solid', color: { kind: 'scheme', value: 'unknown' } },
      { kind: 'solid', color: { kind: 'scheme', value: 'accent1', extra: true } },
    ]) {
      expect(() => normalizeSimpleFill(value, 'Fill')).toThrow(TypeError);
    }
    for (const transparency of [
      '50',
      Number.NaN,
      Number.POSITIVE_INFINITY,
      Number.NEGATIVE_INFINITY,
    ]) {
      expect(() => normalizeSimpleFill({
        kind: 'solid',
        color: { kind: 'srgb', value: 'FFFFFF' },
        transparency,
      }, 'Fill')).toThrow(TypeError);
    }
    for (const transparency of [-0.001, 100.001]) {
      expect(() => normalizeSimpleFill({
        kind: 'solid',
        color: { kind: 'srgb', value: 'FFFFFF' },
        transparency,
      }, 'Fill')).toThrow(RangeError);
    }
  });
});

describe('simple fill choice codec', () => {
  it('renders deterministic none, solid, scheme, and explicit alpha states', () => {
    expect(renderSimpleFill({ kind: 'none' }, 'a:')).toBe('<a:noFill/>');
    expect(renderSimpleFill({
      kind: 'solid',
      color: { kind: 'srgb', value: 'FF0000' },
    }, 'a:')).toBe(
      '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>',
    );
    expect(renderSimpleFill({
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent2' },
      transparency: 50,
    }, 'd:')).toBe(
      '<d:solidFill><d:schemeClr val="accent2"><d:alpha val="50000"/></d:schemeClr></d:solidFill>',
    );
    expect(renderSimpleFill({
      kind: 'solid',
      color: { kind: 'srgb', value: '00FF00' },
      transparency: 0,
    }, 'a:')).toContain('<a:alpha val="100000"/>');
    expect(renderSimpleFill({
      kind: 'solid',
      color: { kind: 'srgb', value: '0000FF' },
      transparency: 100,
    }, 'a:')).toContain('<a:alpha val="0"/>');
  });

  it('reads strict none, sRGB, scheme, and alpha choices with matching prefixes', () => {
    expect(readSimpleFillChoice(parseChoice('<a:noFill/>'), 'a:')).toEqual({ kind: 'none' });
    expect(readSimpleFillChoice(
      parseChoice('<a:solidFill><a:srgbClr val="ff0000"/></a:solidFill>'),
      'a:',
    )).toEqual({ kind: 'solid', color: { kind: 'srgb', value: 'FF0000' } });
    expect(readSimpleFillChoice(
      parseChoice(
        '<d:solidFill><d:schemeClr val="accent3"><d:alpha val="75000"/></d:schemeClr></d:solidFill>',
      ),
      'd:',
    )).toEqual({
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent3' },
      transparency: 25,
    });
    expect(readSimpleFillChoice(
      parseChoice(
        '<a:solidFill><a:srgbClr val="112233"><a:alpha val="49445"/></a:srgbClr></a:solidFill>',
      ),
      'a:',
    )).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: '112233' },
      transparency: 50.555,
    });
    expect(readSimpleFillChoice(
      parseChoice(
        '<a:solidFill><a:srgbClr val="112233"><a:alpha val="100000"/></a:srgbClr></a:solidFill>',
      ),
      'a:',
    )).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: '112233' },
      transparency: 0,
    });
    expect(readSimpleFillChoice(
      parseChoice(
        '<a:solidFill><a:srgbClr val="112233"><a:alpha val="0"/></a:srgbClr></a:solidFill>',
      ),
      'a:',
    )).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: '112233' },
      transparency: 100,
    });
  });

  it('returns undefined for unsupported, malformed, or prefix-mismatched choices', () => {
    const fixtures: Array<[string, string]> = [
      ['<a:noFill custom="x"/>', 'a:'],
      ['<a:noFill><a:ext/></a:noFill>', 'a:'],
      ['<a:noFill/>', 'd:'],
      ['<a:solidFill custom="x"><a:srgbClr val="FF0000"/></a:solidFill>', 'a:'],
      ['<a:solidFill/>', 'a:'],
      ['<a:solidFill><a:srgbClr val="FFF"/></a:solidFill>', 'a:'],
      ['<a:solidFill><a:schemeClr val="unknown"/></a:solidFill>', 'a:'],
      ['<a:solidFill><a:srgbClr x:val="FF0000" xmlns:x="urn:x"/></a:solidFill>', 'a:'],
      ['<a:solidFill><a:srgbClr val="FF0000" custom="x"/></a:solidFill>', 'a:'],
      ['<a:solidFill><a:srgbClr val="FF0000"/><a:schemeClr val="accent1"/></a:solidFill>', 'a:'],
      ['<a:solidFill><a:srgbClr val="FF0000"><a:alpha/></a:srgbClr></a:solidFill>', 'a:'],
      ['<a:solidFill><a:srgbClr val="FF0000"><a:alpha val=""/></a:srgbClr></a:solidFill>', 'a:'],
      ['<a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000" custom="x"/></a:srgbClr></a:solidFill>', 'a:'],
      ['<a:solidFill><a:srgbClr val="FF0000"><a:alpha val="1.5"/></a:srgbClr></a:solidFill>', 'a:'],
      ['<a:solidFill><a:srgbClr val="FF0000"><a:alpha val="-1"/></a:srgbClr></a:solidFill>', 'a:'],
      ['<a:solidFill><a:srgbClr val="FF0000"><a:alpha val="100001"/></a:srgbClr></a:solidFill>', 'a:'],
      ['<a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000"/><a:alpha val="75000"/></a:srgbClr></a:solidFill>', 'a:'],
      ['<a:solidFill><a:srgbClr val="FF0000"><a:tint val="50000"/></a:srgbClr></a:solidFill>', 'a:'],
      ['<a:gradFill/>', 'a:'],
    ];
    for (const [source, prefix] of fixtures) {
      expect(readSimpleFillChoice(parseChoice(source), prefix), source).toBeUndefined();
    }
  });

  it('compares kind, normalized color, and transparency presence exactly', () => {
    expect(simpleFillsEqual({ kind: 'none' }, { kind: 'none' })).toBe(true);
    expect(simpleFillsEqual(undefined, { kind: 'none' })).toBe(false);
    expect(simpleFillsEqual(
      { kind: 'solid', color: { kind: 'srgb', value: 'FF0000' } },
      { kind: 'solid', color: { kind: 'srgb', value: 'FF0000' } },
    )).toBe(true);
    expect(simpleFillsEqual(
      { kind: 'solid', color: { kind: 'srgb', value: 'FF0000' } },
      { kind: 'solid', color: { kind: 'srgb', value: 'FF0000' }, transparency: 0 },
    )).toBe(false);
    expect(simpleFillsEqual(
      { kind: 'solid', color: { kind: 'scheme', value: 'accent1' }, transparency: 25 },
      { kind: 'solid', color: { kind: 'scheme', value: 'accent2' }, transparency: 25 },
    )).toBe(false);
  });
});
