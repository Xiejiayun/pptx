import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import { describe, expect, it } from 'vitest';
import {
  normalizeShapeShadow,
  readSimpleShadow,
  renderSimpleShadow,
  shapeShadowsEqual,
  type NormalizedShapeShadow,
} from './simple-shadow.internal.js';

const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';

function parseShadow(source: string): XmlElement {
  const xml = LosslessXmlDocument.parse(
    `<root xmlns:a="${DRAWING_NAMESPACE}" xmlns:d="${DRAWING_NAMESPACE}" ` +
    `xmlns:x="urn:wrong">${source}</root>`,
  );
  const root = xml.roots[0];
  const shadow = root?.children.find(
    (child): child is XmlElement => child.type === 'element',
  );
  if (!shadow) throw new Error('Fixture has no shadow element');
  return shadow;
}

describe('simple shape shadow normalization', () => {
  it('normalizes outer and inner defaults into detached deep-frozen values', () => {
    const outer = normalizeShapeShadow({ kind: 'outer' }, 'Shape shadow');
    expect(outer).toEqual({
      kind: 'outer',
      color: { kind: 'srgb', value: '000000' },
      opacity: 0.75,
      blur: 8,
      angle: 270,
      distance: 4,
      rotateWithShape: false,
    });
    expect(Object.isFrozen(outer)).toBe(true);
    expect(Object.isFrozen(outer.color)).toBe(true);

    const color = { kind: 'scheme', value: 'accent2' };
    const input = { kind: 'inner', color };
    const inner = normalizeShapeShadow(input, 'Shape shadow');
    color.value = 'accent3';
    input.kind = 'outer';
    expect(inner).toEqual({
      kind: 'inner',
      color: { kind: 'scheme', value: 'accent2' },
      opacity: 0.75,
      blur: 8,
      angle: 270,
      distance: 4,
    });
    expect(inner).not.toBe(input);
    expect(inner.color).not.toBe(color);
    expect(Object.isFrozen(inner)).toBe(true);
    expect(Object.isFrozen(inner.color)).toBe(true);
  });

  it('preserves explicit zero, quantizes every numeric field, and accepts null prototypes', () => {
    expect(normalizeShapeShadow({
      kind: 'outer',
      color: { kind: 'srgb', value: '#a1b2c3' },
      opacity: 0,
      blur: 0,
      angle: 0,
      distance: 0,
      rotateWithShape: true,
    }, 'Shape shadow')).toEqual({
      kind: 'outer',
      color: { kind: 'srgb', value: 'A1B2C3' },
      opacity: 0,
      blur: 0,
      angle: 0,
      distance: 0,
      rotateWithShape: true,
    });

    expect(normalizeShapeShadow({
      kind: 'inner',
      color: { kind: 'scheme', value: 'accent6' },
      opacity: 0.333333,
      blur: 0.333333,
      angle: 123.456789,
      distance: 0.123456,
    }, 'Shape shadow')).toEqual({
      kind: 'inner',
      color: { kind: 'scheme', value: 'accent6' },
      opacity: 33_333 / 100_000,
      blur: 4_233 / 12_700,
      angle: 7_407_407 / 60_000,
      distance: 1_568 / 12_700,
    });

    const color = Object.create(null) as Record<string, unknown>;
    color.kind = 'scheme';
    color.value = 'hlink';
    const shadow = Object.create(null) as Record<string, unknown>;
    shadow.kind = 'outer';
    shadow.color = color;
    shadow.opacity = undefined;
    shadow.blur = undefined;
    shadow.angle = undefined;
    shadow.distance = undefined;
    shadow.rotateWithShape = undefined;
    expect(normalizeShapeShadow(shadow, 'Shape shadow')).toEqual({
      kind: 'outer',
      color: { kind: 'scheme', value: 'hlink' },
      opacity: 0.75,
      blur: 8,
      angle: 270,
      distance: 4,
      rotateWithShape: false,
    });
  });

  it('accepts all numeric boundaries', () => {
    expect(normalizeShapeShadow({
      kind: 'inner',
      opacity: 1,
      blur: 100,
      angle: 359.999,
      distance: 200,
    }, 'Shape shadow')).toMatchObject({
      opacity: 1,
      blur: 100,
      angle: 359.999,
      distance: 200,
    });
  });

  it('rejects non-ordinary, inherited, unknown, symbol, and alias-bearing values', () => {
    class Shadow {
      kind = 'outer';
    }
    for (const value of [null, false, 1, 'shadow', [], new Date(), new Shadow()]) {
      expect(() => normalizeShapeShadow(value, 'Shape shadow')).toThrow(TypeError);
    }
    expect(() => normalizeShapeShadow(Object.create({ kind: 'outer' }), 'Shape shadow'))
      .toThrow(TypeError);
    for (const alias of ['type', 'offset', 'transparency', 'rotate']) {
      expect(() => normalizeShapeShadow({ kind: 'outer', [alias]: undefined }, 'Shape shadow'))
        .toThrow(new RegExp(`unsupported property ${alias}`));
    }
    expect(() => normalizeShapeShadow({ kind: 'outer', [Symbol('unsafe')]: true }, 'Shape shadow'))
      .toThrow(/unsupported property/);
    expect(() => normalizeShapeShadow({
      kind: 'outer',
      color: { kind: 'srgb', value: 'FFFFFF', extra: undefined },
    }, 'Shape shadow')).toThrow(/unsupported property extra/);
  });

  it('rejects accessors without invoking them', () => {
    let calls = 0;
    const kindGetter = Object.defineProperty({}, 'kind', {
      enumerable: true,
      get() {
        calls += 1;
        return 'outer';
      },
    });
    const colorGetter = Object.defineProperty({ kind: 'outer' }, 'color', {
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
    const blurSetter = Object.defineProperty({ kind: 'outer' }, 'blur', {
      enumerable: true,
      set(_value: unknown) {
        calls += 1;
      },
    });
    for (const value of [
      kindGetter,
      colorGetter,
      { kind: 'outer', color: colorValueGetter },
      blurSetter,
    ]) {
      expect(() => normalizeShapeShadow(value, 'Shape shadow')).toThrow(/data property/);
    }
    expect(calls).toBe(0);
  });

  it('rejects invalid kinds, colors, rotate flags, and numeric values', () => {
    for (const value of [
      {},
      { kind: 'none' },
      { kind: 'Outer' },
      { kind: 'shadow' },
      { kind: 1 },
      { kind: 'inner', rotateWithShape: false },
      { kind: 'outer', rotateWithShape: 1 },
      { kind: 'outer', color: null },
      { kind: 'outer', color: { kind: 'rgb', value: 'FFFFFF' } },
      { kind: 'outer', color: { kind: 'srgb', value: 'FFF' } },
      { kind: 'outer', color: { kind: 'srgb', value: 'GGGGGG' } },
      { kind: 'outer', color: { kind: 'scheme', value: 'unknown' } },
    ]) {
      expect(() => normalizeShapeShadow(value, 'Shape shadow')).toThrow(TypeError);
    }

    for (const key of ['opacity', 'blur', 'angle', 'distance'] as const) {
      for (const value of ['1', Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY]) {
        expect(() => normalizeShapeShadow({ kind: 'outer', [key]: value }, 'Shape shadow'))
          .toThrow(TypeError);
      }
    }
    for (const [key, value] of [
      ['opacity', -0.00001],
      ['opacity', 1.00001],
      ['blur', -0.00001],
      ['blur', 100.00001],
      ['angle', -0.00001],
      ['angle', 360],
      ['distance', -0.00001],
      ['distance', 200.00001],
    ] as const) {
      expect(() => normalizeShapeShadow({ kind: 'outer', [key]: value }, 'Shape shadow'))
        .toThrow(RangeError);
    }
  });
});

describe('simple shape shadow codec', () => {
  it('renders deterministic outer and inner shadow XML', () => {
    expect(renderSimpleShadow(
      normalizeShapeShadow({ kind: 'outer' }, 'Shape shadow'),
      'a:',
    )).toBe(
      '<a:outerShdw sx="100000" sy="100000" kx="0" ky="0" algn="bl" ' +
      'rotWithShape="0" blurRad="101600" dist="50800" dir="16200000">' +
      '<a:srgbClr val="000000"><a:alpha val="75000"/></a:srgbClr>' +
      '</a:outerShdw>',
    );
    expect(renderSimpleShadow(
      normalizeShapeShadow({
        kind: 'inner',
        color: { kind: 'scheme', value: 'accent3' },
        opacity: 0,
        blur: 0,
        angle: 0,
        distance: 0,
      }, 'Shape shadow'),
      'd:',
    )).toBe(
      '<d:innerShdw blurRad="0" dist="0" dir="0">' +
      '<d:schemeClr val="accent3"><d:alpha val="0"/></d:schemeClr>' +
      '</d:innerShdw>',
    );
    expect(renderSimpleShadow(
      normalizeShapeShadow({
        kind: 'outer',
        color: { kind: 'srgb', value: 'FFFFFF' },
        opacity: 1,
        blur: 100,
        angle: 359,
        distance: 200,
        rotateWithShape: true,
      }, 'Shape shadow'),
      'a:',
    )).toContain(
      'rotWithShape="1" blurRad="1270000" dist="2540000" dir="21540000">' +
      '<a:srgbClr val="FFFFFF"><a:alpha val="100000"/></a:srgbClr>',
    );
  });

  it('reads canonical, alternate-prefix, and optional-default states', () => {
    const outer = readSimpleShadow(parseShadow(
      '<a:outerShdw sx="100000" sy="100000" kx="0" ky="0" algn="bl" ' +
      'rotWithShape="0" blurRad="92075" dist="69850" dir="7404000">' +
      '<a:srgbClr val="123abc"><a:alpha val="42000"/></a:srgbClr>' +
      '</a:outerShdw>',
    ), 'a:');
    expect(outer).toEqual({
      kind: 'outer',
      color: { kind: 'srgb', value: '123ABC' },
      opacity: 0.42,
      blur: 7.25,
      angle: 123.4,
      distance: 5.5,
      rotateWithShape: false,
    });
    expect(Object.isFrozen(outer)).toBe(true);
    expect(Object.isFrozen(outer?.color)).toBe(true);

    const inner = readSimpleShadow(parseShadow(
      '<d:innerShdw blurRad="1270000" dist="2540000" dir="21599940">' +
      '<d:schemeClr val="accent5"/></d:innerShdw>',
    ), 'd:');
    expect(inner).toEqual({
      kind: 'inner',
      color: { kind: 'scheme', value: 'accent5' },
      opacity: 1,
      blur: 100,
      angle: 359.999,
      distance: 200,
    });

    expect(readSimpleShadow(parseShadow(
      '<a:outerShdw><a:srgbClr val="000000"/></a:outerShdw>',
    ), 'a:')).toEqual({
      kind: 'outer',
      color: { kind: 'srgb', value: '000000' },
      opacity: 1,
      blur: 0,
      angle: 0,
      distance: 0,
      rotateWithShape: true,
    });
    for (const lexical of ['0', 'false', '1', 'true']) {
      expect(readSimpleShadow(parseShadow(
        `<a:outerShdw rotWithShape="${lexical}">` +
        '<a:srgbClr val="000000"/></a:outerShdw>',
      ), 'a:')).toMatchObject({
        rotateWithShape: lexical === '1' || lexical === 'true',
      });
    }
  });

  it('returns fresh detached snapshots without mutating source state', () => {
    const element = parseShadow(
      '<a:innerShdw blurRad="12700" dist="25400" dir="60000">' +
      '<a:srgbClr val="ABCDEF"><a:alpha val="50000"/></a:srgbClr>' +
      '</a:innerShdw>',
    );
    const first = readSimpleShadow(element, 'a:');
    const second = readSimpleShadow(element, 'a:');
    expect(first).toEqual(second);
    expect(first).not.toBe(second);
    expect(first?.color).not.toBe(second?.color);
    expect(Object.isFrozen(first)).toBe(true);
    expect(Object.isFrozen(first?.color)).toBe(true);
  });

  it('returns undefined for malformed attributes, ranges, colors, transforms, and namespaces', () => {
    const fixtures: Array<[string, string]> = [
      ['<a:glow><a:srgbClr val="000000"/></a:glow>', 'a:'],
      ['<x:outerShdw><x:srgbClr val="000000"/></x:outerShdw>', 'a:'],
      ['<a:outerShdw xmlns:a="urn:wrong"><a:srgbClr val="000000"/></a:outerShdw>', 'a:'],
      ['<a:outerShdw custom="1"><a:srgbClr val="000000"/></a:outerShdw>', 'a:'],
      ['<a:outerShdw x:blurRad="1"><a:srgbClr val="000000"/></a:outerShdw>', 'a:'],
      ['<a:innerShdw rotWithShape="0"><a:srgbClr val="000000"/></a:innerShdw>', 'a:'],
      ['<a:innerShdw sx="100000"><a:srgbClr val="000000"/></a:innerShdw>', 'a:'],
      ['<a:outerShdw sx="100000"><a:srgbClr val="000000"/></a:outerShdw>', 'a:'],
      ['<a:outerShdw sx="99999" sy="100000" kx="0" ky="0" algn="bl">' +
        '<a:srgbClr val="000000"/></a:outerShdw>', 'a:'],
      ['<a:outerShdw sx="100000" sy="100000" kx="0" ky="0" algn="ctr">' +
        '<a:srgbClr val="000000"/></a:outerShdw>', 'a:'],
      ['<a:outerShdw blurRad=""><a:srgbClr val="000000"/></a:outerShdw>', 'a:'],
      ['<a:outerShdw blurRad="1.5"><a:srgbClr val="000000"/></a:outerShdw>', 'a:'],
      ['<a:outerShdw blurRad="-1"><a:srgbClr val="000000"/></a:outerShdw>', 'a:'],
      ['<a:outerShdw blurRad="1270001"><a:srgbClr val="000000"/></a:outerShdw>', 'a:'],
      ['<a:outerShdw dist="2540001"><a:srgbClr val="000000"/></a:outerShdw>', 'a:'],
      ['<a:outerShdw dir="21600000"><a:srgbClr val="000000"/></a:outerShdw>', 'a:'],
      ['<a:outerShdw rotWithShape="yes"><a:srgbClr val="000000"/></a:outerShdw>', 'a:'],
      ['<a:outerShdw/>', 'a:'],
      ['<a:outerShdw><a:srgbClr val="000000"/><a:schemeClr val="accent1"/></a:outerShdw>', 'a:'],
      ['<a:outerShdw><x:srgbClr val="000000"/></a:outerShdw>', 'a:'],
      ['<a:outerShdw><a:prstClr val="black"/></a:outerShdw>', 'a:'],
      ['<a:outerShdw><a:sysClr val="windowText"/></a:outerShdw>', 'a:'],
      ['<a:outerShdw><a:srgbClr/></a:outerShdw>', 'a:'],
      ['<a:outerShdw><a:srgbClr val="FFF"/></a:outerShdw>', 'a:'],
      ['<a:outerShdw><a:srgbClr x:val="000000"/></a:outerShdw>', 'a:'],
      ['<a:outerShdw><a:srgbClr val="000000" custom="1"/></a:outerShdw>', 'a:'],
      ['<a:outerShdw><a:schemeClr val="unknown"/></a:outerShdw>', 'a:'],
      ['<a:outerShdw><a:srgbClr val="000000"><a:alpha/></a:srgbClr></a:outerShdw>', 'a:'],
      ['<a:outerShdw><a:srgbClr val="000000"><a:alpha val="1.5"/></a:srgbClr></a:outerShdw>', 'a:'],
      ['<a:outerShdw><a:srgbClr val="000000"><a:alpha val="-1"/></a:srgbClr></a:outerShdw>', 'a:'],
      ['<a:outerShdw><a:srgbClr val="000000"><a:alpha val="100001"/></a:srgbClr></a:outerShdw>', 'a:'],
      ['<a:outerShdw><a:srgbClr val="000000"><a:alpha val="50000" custom="1"/></a:srgbClr></a:outerShdw>', 'a:'],
      ['<a:outerShdw><a:srgbClr val="000000"><x:alpha val="50000"/></a:srgbClr></a:outerShdw>', 'a:'],
      ['<a:outerShdw><a:srgbClr val="000000"><a:alpha val="50000"><a:ext/></a:alpha></a:srgbClr></a:outerShdw>', 'a:'],
      ['<a:outerShdw><a:srgbClr val="000000"><a:alpha val="50000"/><a:alpha val="75000"/></a:srgbClr></a:outerShdw>', 'a:'],
      ['<a:outerShdw><a:srgbClr val="000000"><a:tint val="50000"/></a:srgbClr></a:outerShdw>', 'a:'],
      ['<a:outerShdw><a:srgbClr val="000000"/></a:outerShdw>', 'd:'],
    ];
    for (const [source, prefix] of fixtures) {
      expect(readSimpleShadow(parseShadow(source), prefix), source).toBeUndefined();
    }
  });

  it('compares every normalized field', () => {
    const outer = normalizeShapeShadow({ kind: 'outer' }, 'Shape shadow');
    if (outer.kind !== 'outer') throw new Error('Expected outer shadow');
    expect(shapeShadowsEqual(undefined, undefined)).toBe(true);
    expect(shapeShadowsEqual(outer, outer)).toBe(true);
    expect(shapeShadowsEqual(outer, {
      ...outer,
      color: { ...outer.color },
    })).toBe(true);
    expect(shapeShadowsEqual(undefined, outer)).toBe(false);
    expect(shapeShadowsEqual(outer, normalizeShapeShadow({ kind: 'inner' }, 'Shape shadow')))
      .toBe(false);

    const variations: NormalizedShapeShadow[] = [
      { ...outer, color: { kind: 'srgb', value: 'FFFFFF' } },
      { ...outer, color: { kind: 'scheme', value: 'tx1' } },
      { ...outer, opacity: 0.5 },
      { ...outer, blur: 9 },
      { ...outer, angle: 90 },
      { ...outer, distance: 5 },
      { ...outer, rotateWithShape: true },
    ];
    for (const variation of variations) {
      expect(shapeShadowsEqual(outer, variation)).toBe(false);
    }
  });
});
