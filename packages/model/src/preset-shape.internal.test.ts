import { describe, expect, it } from 'vitest';
import { createHash } from 'node:crypto';
import { LosslessXmlDocument } from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
import {
  PRESET_SHAPE_TYPES,
  type ShapeAdjustment,
  type ShapeArrowType,
  type ShapeLineDash,
  type ShapeShadow,
} from './preset-shape.js';
import type { ShapeArrowTypeValue } from './shape-arrows.internal.js';
import type { SimpleLineDash } from './simple-line.internal.js';
import type { CustomGeometry } from './custom-geometry.js';
import {
  normalizeCustomShape,
  normalizePresetShape,
  readPresetShapeType,
  renderCustomShapeXml,
  renderPresetShapeXml,
  replacePresetShapeType,
} from './preset-shape.internal.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
type PublicDashMatchesInternal =
  [ShapeLineDash] extends [SimpleLineDash]
    ? [SimpleLineDash] extends [ShapeLineDash]
      ? true
      : false
    : false;
type PublicArrowMatchesInternal =
  [ShapeArrowType] extends [ShapeArrowTypeValue]
    ? [ShapeArrowTypeValue] extends [ShapeArrowType]
      ? true
      : false
    : false;

const customGeometry: CustomGeometry = {
  paths: [{
    width: 3_657_600,
    height: 2_743_200,
    commands: [
      { kind: 'moveTo', point: { x: 0, y: 0 } },
      { kind: 'lineTo', point: { x: 3_657_600, y: 0 } },
      { kind: 'lineTo', point: { x: 1_828_800, y: 2_743_200 } },
      { kind: 'close' },
    ],
  }],
};

function parseShape(source: string) {
  const xml = LosslessXmlDocument.parse(source);
  const shape = xml.roots[0];
  if (!shape) throw new Error('Fixture has no shape root');
  return { xml, shape };
}

function shapeFixture(
  geometry = '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>',
  options: { readonly rootName?: string; readonly namespace?: string; readonly properties?: string } = {},
): string {
  const rootName = options.rootName ?? 'p:sp';
  const namespace = options.namespace ?? PRESENTATION_NAMESPACE;
  const properties = options.properties ?? `<p:spPr>${geometry}<a:noFill/><a:ln/></p:spPr>`;
  return `<${rootName} xmlns:p="${namespace}" xmlns:a="${DRAWING_NAMESPACE}">` +
    '<p:nvSpPr><p:cNvPr id="7" name="Keep"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
    `${properties}<p:txBody><a:bodyPr/><a:p><a:r><a:t>KEEP</a:t></a:r></a:p></p:txBody>` +
    `</${rootName}>`;
}

describe('preset shape catalog', () => {
  it('publishes the exact frozen canonical PptxGenJS 4.0.1 shape set', () => {
    expect(PRESET_SHAPE_TYPES).toHaveLength(178);
    expect(new Set(PRESET_SHAPE_TYPES)).toHaveLength(178);
    expect(Object.isFrozen(PRESET_SHAPE_TYPES)).toBe(true);
    expect(PRESET_SHAPE_TYPES[0]).toBe('accentBorderCallout1');
    expect(PRESET_SHAPE_TYPES.at(-1)).toBe('wedgeRoundRectCallout');
    expect(PRESET_SHAPE_TYPES).toContain('foldedCorner');
    expect(PRESET_SHAPE_TYPES).not.toContain('folderCorner' as never);
    expect(PRESET_SHAPE_TYPES).not.toContain('custGeom' as never);
    expect(createHash('sha256').update(PRESET_SHAPE_TYPES.join('\n')).digest('hex')).toBe(
      '4b2d864583049a8eb02457e93504f608b76d8e04723e45eb4ec62b1bbb129c3d',
    );

    expect(() => {
      (PRESET_SHAPE_TYPES as unknown as string[]).push('rect');
    }).toThrow(TypeError);
    expect(PRESET_SHAPE_TYPES).toHaveLength(178);
  });
});

describe('normalizePresetShape', () => {
  it('uses one-inch defaults and detaches rounded primitive values', () => {
    expect(normalizePresetShape('rect', undefined)).toEqual({
      type: 'rect',
      name: undefined,
      fill: { kind: 'none' },
      line: undefined,
      arrows: undefined,
      hyperlink: undefined,
      shadow: undefined,
      adjustments: [],
      x: 914_400,
      y: 914_400,
      width: 914_400,
      height: 914_400,
      rotation: 0,
      flipHorizontal: false,
      flipVertical: false,
    });

    const options = Object.create(null) as Record<string, unknown>;
    options.name = 'A & <B>';
    options.x = 1.4;
    options.y = -1.5;
    options.width = 2.6;
    options.height = 3.5;
    options.rotation = 2.5;
    options.flipHorizontal = true;
    options.flipVertical = true;
    const color = { kind: 'srgb', value: '#ff0000' };
    const fill = { kind: 'solid', color, transparency: 33.3334 };
    options.fill = fill;
    const normalized = normalizePresetShape('ellipse', options);
    options.name = 'Changed';
    options.x = 99;
    color.value = '000000';
    fill.transparency = 1;

    expect(normalized).toEqual({
      type: 'ellipse',
      name: 'A & <B>',
      fill: {
        kind: 'solid',
        color: { kind: 'srgb', value: 'FF0000' },
        transparency: 33.333,
      },
      line: undefined,
      arrows: undefined,
      hyperlink: undefined,
      shadow: undefined,
      adjustments: [],
      x: 1,
      y: -1,
      width: 3,
      height: 4,
      rotation: 3,
      flipHorizontal: true,
      flipVertical: true,
    });
    expect(Object.isFrozen(normalized)).toBe(true);
  });

  it('normalizes detached preset shape lines with exact width and dash defaults', () => {
    const color = { kind: 'srgb', value: '#123456' };
    const line = {
      kind: 'line' as const,
      color,
      transparency: 12.3456,
      width: 2.50001,
      dash: 'dashDot' as const,
    };
    const normalized = normalizePresetShape('roundRect', { line });
    color.value = 'FFFFFF';
    line.transparency = 90;
    line.width = 9;

    expect(normalized.line).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: '123456' },
      transparency: 12.346,
      width: 31_750 / 12_700,
      dash: 'dashDot',
    });
    expect(normalized.line).not.toBe(line);
    if (normalized.line?.kind === 'line') expect(normalized.line.color).not.toBe(color);

    expect(normalizePresetShape('ellipse', {
      line: { kind: 'line', color: { kind: 'scheme', value: 'accent3' } },
    }).line).toEqual({
      kind: 'line',
      color: { kind: 'scheme', value: 'accent3' },
      width: 1,
      dash: 'solid',
    });
    expect(normalizePresetShape('star5', { line: { kind: 'none' } }).line)
      .toEqual({ kind: 'none' });

    const publicDash: ShapeLineDash = 'lgDashDotDot';
    const internalDash: SimpleLineDash = publicDash;
    const publicAgain: ShapeLineDash = internalDash;
    const exactDashUnion: PublicDashMatchesInternal = true;
    expect(publicAgain).toBe('lgDashDotDot');
    expect(exactDashUnion).toBe(true);
  });

  it('normalizes detached preset shape arrows with an exact public token union', () => {
    const arrows: { begin: ShapeArrowType; end: ShapeArrowType | undefined } = {
      begin: 'triangle',
      end: 'arrow',
    };
    const normalized = normalizePresetShape('line', { arrows });
    arrows.begin = 'diamond';
    arrows.end = undefined;

    expect(normalized.arrows).toEqual({ begin: 'triangle', end: 'arrow' });
    expect(normalized.arrows).not.toBe(arrows);
    expect(Object.isFrozen(normalized.arrows)).toBe(true);
    expect(normalizePresetShape('lineInv', { arrows: {} }).arrows).toEqual({});
    expect(normalizePresetShape('line', {
      arrows: { begin: undefined, end: 'none' },
    }).arrows).toEqual({ end: 'none' });

    const publicType: ShapeArrowType = 'stealth';
    const internalType: ShapeArrowTypeValue = publicType;
    const publicAgain: ShapeArrowType = internalType;
    const exactArrowUnion: PublicArrowMatchesInternal = true;
    expect(publicAgain).toBe('stealth');
    expect(exactArrowUnion).toBe(true);
  });

  it('normalizes detached URL and slide hyperlinks with direct tooltip state', () => {
    const hyperlink: { url: string; tooltip?: string } = {
      url: 'https://example.com?a=1&b=2',
      tooltip: 'Visit & learn',
    };
    const normalized = normalizePresetShape('rect', { hyperlink });
    hyperlink.url = 'https://changed.example';
    delete hyperlink.tooltip;

    expect(normalized.hyperlink).toEqual({
      url: 'https://example.com?a=1&b=2',
      tooltip: 'Visit & learn',
    });
    expect(normalized.hyperlink).not.toBe(hyperlink);
    expect(Object.isFrozen(normalized.hyperlink)).toBe(true);
    expect(normalizePresetShape('actionButtonForwardNext', {
      hyperlink: { slide: 2, tooltip: '' },
    }).hyperlink).toEqual({ slide: 2, tooltip: '' });
    expect(normalizePresetShape('actionButtonHome', {
      hyperlink: undefined,
    }).hyperlink).toBeUndefined();
  });

  it('normalizes detached preset shape shadows with exact defaults and explicit zeros', () => {
    expect(normalizePresetShape('rect', { shadow: { kind: 'outer' } }).shadow).toEqual({
      kind: 'outer',
      color: { kind: 'srgb', value: '000000' },
      opacity: 0.75,
      blur: 8,
      angle: 270,
      distance: 4,
      rotateWithShape: false,
    });
    const color: { kind: 'srgb'; value: string } = { kind: 'srgb', value: '#123abc' };
    const shadow: ShapeShadow = {
      kind: 'outer',
      color,
      opacity: 0.42,
      blur: 7.25,
      angle: 123.4,
      distance: 5.5,
      rotateWithShape: true,
    };
    const normalized = normalizePresetShape('roundRect', { shadow });
    color.value = 'FFFFFF';

    expect(normalized.shadow).toEqual({
      kind: 'outer',
      color: { kind: 'srgb', value: '123ABC' },
      opacity: 0.42,
      blur: 7.25,
      angle: 123.4,
      distance: 5.5,
      rotateWithShape: true,
    });
    expect(normalized.shadow).not.toBe(shadow);
    expect(Object.isFrozen(normalized.shadow)).toBe(true);
    expect(Object.isFrozen(normalized.shadow?.color)).toBe(true);
    expect(normalizePresetShape('ellipse', {
      shadow: {
        kind: 'inner',
        color: { kind: 'scheme', value: 'accent2' },
        opacity: 0,
        blur: 0,
        angle: 0,
        distance: 0,
      },
    }).shadow).toEqual({
      kind: 'inner',
      color: { kind: 'scheme', value: 'accent2' },
      opacity: 0,
      blur: 0,
      angle: 0,
      distance: 0,
    });
    expect(normalizePresetShape('star5', { shadow: undefined }).shadow).toBeUndefined();
  });

  it('normalizes detached ordered preset shape adjustments and preserves explicit zero', () => {
    const adjustments: { name: string; value: number }[] = [
      { name: 'adj1', value: 16_200_000 },
      { name: 'adj2', value: 0 },
      { name: 'adj3', value: 25_000 },
    ];
    const normalized = normalizePresetShape('blockArc', { adjustments });
    adjustments[0]!.name = 'changed';
    adjustments[0]!.value = 1;

    expect(normalized.adjustments).toEqual([
      { name: 'adj1', value: 16_200_000 },
      { name: 'adj2', value: 0 },
      { name: 'adj3', value: 25_000 },
    ]);
    expect(Object.isFrozen(normalized.adjustments)).toBe(true);
    expect(normalized.adjustments.every(Object.isFrozen)).toBe(true);
    expect(normalizePresetShape('rect', { adjustments: [] }).adjustments).toEqual([]);
    expect(normalizePresetShape('rect', { adjustments: undefined } as never).adjustments)
      .toEqual([]);

    const publicValue: readonly ShapeAdjustment[] = normalized.adjustments;
    expect(publicValue[1]?.value).toBe(0);
  });

  it('rejects unsafe hyperlink values and accessors without invoking them', () => {
    let calls = 0;
    const hyperlinkGetter = Object.defineProperty({}, 'url', {
      enumerable: true,
      get() {
        calls += 1;
        return 'https://example.com';
      },
    });
    for (const hyperlink of [
      null,
      'https://example.com',
      {},
      { url: 'https://example.com', slide: 1 },
      { url: '' },
      { slide: 0 },
      { slide: 1.5 },
      { url: 'https://example.com', target: '_blank' },
      { url: 'https://example.com', [Symbol('unsafe')]: true },
      hyperlinkGetter,
    ]) {
      expect(() => normalizePresetShape('rect', { hyperlink })).toThrow();
    }
    expect(calls).toBe(0);
  });

  it('rejects unknown types and non-ordinary option containers', () => {
    for (const type of [undefined, null, 7, '', 'folderCorner', 'custGeom', 'RECT']) {
      expect(() => normalizePresetShape(type, undefined)).toThrow(TypeError);
    }
    class Options {
      x = 1;
    }
    for (const options of [null, false, 1, 'x', [], new Date(), new Options()]) {
      expect(() => normalizePresetShape('rect', options)).toThrow(TypeError);
    }
    const inherited = Object.create({ x: 1 });
    expect(() => normalizePresetShape('rect', inherited)).toThrow(TypeError);
  });

  it('rejects unsupported, symbolic, and accessor properties without invoking accessors', () => {
    expect(() => normalizePresetShape('rect', { opacity: 1 })).toThrow(/unsupported property opacity/);
    expect(() => normalizePresetShape('rect', { [Symbol('unsafe')]: true })).toThrow(/unsupported property/);

    let calls = 0;
    const getter = Object.defineProperty({}, 'x', {
      enumerable: true,
      get() {
        calls += 1;
        return 1;
      },
    });
    const setter = Object.defineProperty({}, 'y', {
      enumerable: true,
      set(_value: unknown) {
        calls += 1;
      },
    });
    expect(() => normalizePresetShape('rect', getter)).toThrow(/data property/);
    expect(() => normalizePresetShape('rect', setter)).toThrow(/data property/);
    expect(calls).toBe(0);
  });

  it('rejects unsafe names, coordinates, extents, rotation, and flips', () => {
    for (const name of [1, false, 'bad\u0000xml']) {
      expect(() => normalizePresetShape('rect', { name })).toThrow(TypeError);
    }
    for (const key of ['x', 'y', 'width', 'height', 'rotation'] as const) {
      for (const value of ['1', Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY]) {
        expect(() => normalizePresetShape('rect', { [key]: value })).toThrow(TypeError);
      }
      for (const value of [Number.MAX_SAFE_INTEGER + 1, Number.MIN_SAFE_INTEGER - 1]) {
        expect(() => normalizePresetShape('rect', { [key]: value })).toThrow(RangeError);
      }
    }
    for (const key of ['width', 'height'] as const) {
      for (const value of [0, -1, 0.49]) {
        expect(() => normalizePresetShape('rect', { [key]: value })).toThrow(RangeError);
      }
    }
    for (const rotation of [-21_600_001, 21_600_001]) {
      expect(() => normalizePresetShape('rect', { rotation })).toThrow(RangeError);
    }
    for (const key of ['flipHorizontal', 'flipVertical'] as const) {
      for (const value of [0, 1, '', 'false', null]) {
        expect(() => normalizePresetShape('rect', { [key]: value })).toThrow(TypeError);
      }
    }
  });

  it('rejects invalid preset shape fills', () => {
    for (const fill of [
      null,
      [],
      { kind: 'gradient' },
      { kind: 'solid' },
      { kind: 'none', color: undefined },
      { kind: 'solid', color: { kind: 'srgb', value: 'FFF' } },
      { kind: 'solid', color: { kind: 'scheme', value: 'unknown' } },
      {
        kind: 'solid',
        color: { kind: 'srgb', value: 'FFFFFF' },
        transparency: Number.NaN,
      },
      {
        kind: 'solid',
        color: { kind: 'srgb', value: 'FFFFFF' },
        transparency: 101,
      },
      {
        kind: 'solid',
        color: { kind: 'srgb', value: 'FFFFFF' },
        alpha: 40,
      },
      { kind: 'none', type: 'none' },
    ]) {
      expect(() => normalizePresetShape('rect', { fill }), JSON.stringify(fill)).toThrow();
    }
  });

  it('rejects invalid preset shape lines and PptxGenJS aliases', () => {
    for (const line of [
      null,
      [],
      { kind: 'solid' },
      { kind: 'line' },
      { kind: 'none', width: undefined },
      { kind: 'line', color: { kind: 'srgb', value: 'FFF' } },
      { kind: 'line', color: { kind: 'scheme', value: 'unknown' } },
      {
        kind: 'line',
        color: { kind: 'srgb', value: 'FFFFFF' },
        transparency: 101,
      },
      {
        kind: 'line',
        color: { kind: 'srgb', value: 'FFFFFF' },
        width: 1_584.001,
      },
      {
        kind: 'line',
        color: { kind: 'srgb', value: 'FFFFFF' },
        dash: 'dot',
      },
      { type: 'none' },
      { color: 'FFFFFF', dashType: 'dash' },
      { kind: 'line', color: { kind: 'srgb', value: 'FFFFFF' }, alpha: 40 },
      { kind: 'line', color: { kind: 'srgb', value: 'FFFFFF' }, lineDash: 'dash' },
    ]) {
      expect(() => normalizePresetShape('rect', { line }), JSON.stringify(line)).toThrow();
    }
  });

  it('rejects invalid preset shape shadows and PptxGenJS aliases without invoking accessors', () => {
    let calls = 0;
    const shadowGetter = Object.defineProperty({}, 'kind', {
      enumerable: true,
      get() {
        calls += 1;
        return 'outer';
      },
    });
    for (const shadow of [
      null,
      [],
      {},
      { kind: 'none' },
      { kind: 'Outer' },
      { kind: 'inner', rotateWithShape: false },
      { kind: 'outer', opacity: -1 },
      { kind: 'outer', blur: 101 },
      { kind: 'outer', angle: 360 },
      { kind: 'outer', distance: 201 },
      { kind: 'outer', color: { kind: 'srgb', value: 'FFF' } },
      { kind: 'outer', color: { kind: 'scheme', value: 'unknown' } },
      { kind: 'outer', type: 'outer' },
      { kind: 'outer', offset: 4 },
      { kind: 'outer', [Symbol('unsafe')]: true },
      shadowGetter,
    ]) {
      expect(() => normalizePresetShape('rect', { shadow }), String(shadow)).toThrow();
    }
    expect(calls).toBe(0);
  });

  it('rejects invalid preset shape arrows and PptxGenJS aliases', () => {
    class Arrows {
      begin = 'arrow';
    }
    for (const arrows of [
      null,
      false,
      [],
      new Date(),
      new Arrows(),
      Object.create({ begin: 'arrow' }),
      { begin: null },
      { begin: '' },
      { begin: 'Arrow' },
      { begin: 'bogus' },
      { end: 'TRIANGLE' },
      { beginArrowType: 'arrow' },
      { endArrowType: 'arrow' },
      { lineHead: 'arrow' },
      { lineTail: 'arrow' },
      { begin: 'arrow', [Symbol('unsafe')]: true },
    ]) {
      expect(() => normalizePresetShape('line', { arrows }), JSON.stringify(arrows)).toThrow();
    }
  });
});

describe('normalizeCustomShape', () => {
  it('shares detached defaults and every styled option with preset creation', () => {
    expect(normalizeCustomShape(customGeometry)).toEqual({
      geometry: customGeometry,
      name: undefined,
      fill: { kind: 'none' },
      line: undefined,
      arrows: undefined,
      hyperlink: undefined,
      shadow: undefined,
      x: 914_400,
      y: 914_400,
      width: 914_400,
      height: 914_400,
      rotation: 0,
      flipHorizontal: false,
      flipVertical: false,
    });

    const mutableGeometry = structuredClone(customGeometry) as {
      paths: Array<{
        width: number;
        height: number;
        commands: Array<CustomGeometry['paths'][number]['commands'][number]>;
      }>;
    };
    const options = {
      name: 'A & <Custom>',
      fill: { kind: 'solid' as const, color: { kind: 'scheme' as const, value: 'accent2' as const } },
      line: {
        kind: 'line' as const,
        color: { kind: 'srgb' as const, value: '123ABC' },
        width: 2,
      },
      arrows: { begin: 'stealth' as const, end: 'triangle' as const },
      hyperlink: { url: 'https://example.com/custom', tooltip: 'Custom' },
      shadow: { kind: 'outer' as const },
      x: 1.4,
      y: -1.5,
      width: 3.6,
      height: 4.5,
      rotation: 2.5,
      flipHorizontal: true,
      flipVertical: true,
    };
    const normalized = normalizeCustomShape(mutableGeometry, options);
    const { type: _type, adjustments: _adjustments, ...presetOptions } =
      normalizePresetShape('rect', options);
    expect(normalized).toEqual({ geometry: customGeometry, ...presetOptions });
    expect(Object.isFrozen(normalized)).toBe(true);
    expect(Object.isFrozen(normalized.geometry.paths[0]?.commands)).toBe(true);

    mutableGeometry.paths[0]!.width = 1;
    mutableGeometry.paths[0]!.commands.splice(0);
    options.name = 'Changed';
    options.line.color.value = 'FFFFFF';
    expect(normalized.geometry).toEqual(customGeometry);
    expect(normalized.name).toBe('A & <Custom>');
    expect(normalized.line).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: '123ABC' },
      transparency: undefined,
      width: 2,
      dash: 'solid',
    });
  });

  it('rejects preset-only, unknown, symbol, and accessor options without invoking getters', () => {
    expect(() => normalizeCustomShape(customGeometry, { adjustments: undefined } as never))
      .toThrow(/unsupported property adjustments/);
    expect(() => normalizeCustomShape(customGeometry, { unknown: true } as never))
      .toThrow(/unsupported property unknown/);
    expect(() => normalizeCustomShape(customGeometry, { [Symbol('unsafe')]: true } as never))
      .toThrow(/unsupported property/);
    let calls = 0;
    const accessor = Object.defineProperty({}, 'name', {
      enumerable: true,
      get() {
        calls += 1;
        return 'Unsafe';
      },
    });
    expect(() => normalizeCustomShape(customGeometry, accessor)).toThrow(/data property/);
    expect(calls).toBe(0);
  });
});

describe('preset shape XML codec', () => {
  it('renders custom geometry through the byte-identical shared shape skeleton', () => {
    const options = {
      name: 'Custom triangle',
      fill: { kind: 'solid' as const, color: { kind: 'scheme' as const, value: 'accent1' as const } },
      line: {
        kind: 'line' as const,
        color: { kind: 'srgb' as const, value: '112233' },
        width: 2,
      },
      arrows: { end: 'triangle' as const },
      shadow: { kind: 'outer' as const },
      hyperlink: { url: 'https://example.com/custom' },
      x: 914_400,
      y: 1_828_800,
      width: 3_657_600,
      height: 2_743_200,
      rotation: 2_700_000,
      flipHorizontal: true,
    };
    const preset = renderPresetShapeXml(7, normalizePresetShape('rect', options), 'rId7');
    const custom = renderCustomShapeXml(7, normalizeCustomShape(customGeometry, options), 'rId7');
    expect(custom).toBe(preset.replace(
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>',
      '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>' +
      '<a:rect l="l" t="t" r="r" b="b"/><a:pathLst>' +
      '<a:path w="3657600" h="2743200">' +
      '<a:moveTo><a:pt x="0" y="0"/></a:moveTo>' +
      '<a:lnTo><a:pt x="3657600" y="0"/></a:lnTo>' +
      '<a:lnTo><a:pt x="1828800" y="2743200"/></a:lnTo>' +
      '<a:close/></a:path></a:pathLst></a:custGeom>',
    ));
    expect(() => renderCustomShapeXml(
      8,
      normalizeCustomShape(customGeometry, { hyperlink: { url: 'https://example.com' } }),
    )).toThrow(TypeError);
    expect(() => renderCustomShapeXml(
      9,
      normalizeCustomShape(customGeometry),
      'rId9',
    )).toThrow(TypeError);
  });

  it('renders the exact deterministic default skeleton', () => {
    const omitted = renderPresetShapeXml(2, normalizePresetShape('rect', undefined));
    expect(omitted).toBe(
      `<p:sp xmlns:p="${PRESENTATION_NAMESPACE}" xmlns:a="${DRAWING_NAMESPACE}">` +
      '<p:nvSpPr><p:cNvPr id="2" name="Shape 2"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="914400" y="914400"/>' +
      '<a:ext cx="914400" cy="914400"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln/></p:spPr></p:sp>',
    );
    expect(renderPresetShapeXml(
      2,
      normalizePresetShape('rect', { shadow: undefined }),
    )).toBe(omitted);
  });

  it('renders exact URL and internal hyperlinks and requires a paired relationship ID', () => {
    const external = renderPresetShapeXml(7, normalizePresetShape('rect', {
      hyperlink: {
        url: 'https://example.com?a=1&b=2',
        tooltip: 'Visit & learn',
      },
    }), 'rId7');
    expect(external).toContain(
      `<p:cNvPr id="7" name="Shape 7"><a:hlinkClick ` +
      `xmlns:r="${RELATIONSHIP_NAMESPACE}" r:id="rId7" ` +
      'tooltip="Visit &amp; learn"/></p:cNvPr>',
    );
    expect(external).not.toContain('ppaction://hlinksldjump');

    const internal = renderPresetShapeXml(
      8,
      normalizePresetShape('actionButtonForwardNext', {
        hyperlink: { slide: 2, tooltip: '' },
      }),
      'rId8',
    );
    expect(internal).toContain(
      'r:id="rId8" tooltip="" action="ppaction://hlinksldjump"',
    );
    expect(internal).toContain('<p:cNvPr id="8" name="Shape 8">');

    expect(() => renderPresetShapeXml(
      9,
      normalizePresetShape('rect', { hyperlink: { url: 'https://example.com' } }),
    )).toThrow(TypeError);
    expect(() => renderPresetShapeXml(
      10,
      normalizePresetShape('rect'),
      'rId10',
    )).toThrow(TypeError);
  });

  it('renders escaped names and only non-default transform attributes', () => {
    const rendered = renderPresetShapeXml(7, normalizePresetShape('lineInv', {
      name: 'A & <"B">',
      x: -3,
      y: 4,
      width: 5,
      height: 6,
      rotation: -2,
      flipHorizontal: true,
      flipVertical: true,
    }));
    expect(rendered).toContain('id="7" name="A &amp; &lt;&quot;B&quot;&gt;"');
    expect(rendered).toContain('<a:xfrm rot="-2" flipH="1" flipV="1">');
    expect(rendered).toContain('<a:off x="-3" y="4"/><a:ext cx="5" cy="6"/>');
    expect(rendered).toContain('<a:prstGeom prst="lineInv"><a:avLst/></a:prstGeom>');
  });

  it('renders strict shape fills immediately after geometry and before the empty line', () => {
    expect(renderPresetShapeXml(7, normalizePresetShape('rect', {
      fill: {
        kind: 'solid',
        color: { kind: 'srgb', value: '#ff0000' },
      },
    }))).toContain(
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
      '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ln/>',
    );
    expect(renderPresetShapeXml(8, normalizePresetShape('ellipse', {
      fill: {
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent2' },
        transparency: 25,
      },
    }))).toContain(
      '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>' +
      '<a:solidFill><a:schemeClr val="accent2"><a:alpha val="75000"/>' +
      '</a:schemeClr></a:solidFill><a:ln/>',
    );
    expect(renderPresetShapeXml(9, normalizePresetShape('star5', {
      fill: {
        kind: 'solid',
        color: { kind: 'srgb', value: '00FF00' },
        transparency: 0,
      },
    }))).toContain(
      '<a:prstGeom prst="star5"><a:avLst/></a:prstGeom>' +
      '<a:solidFill><a:srgbClr val="00FF00"><a:alpha val="100000"/>' +
      '</a:srgbClr></a:solidFill><a:ln/>',
    );
    expect(renderPresetShapeXml(10, normalizePresetShape('diamond', {
      fill: { kind: 'none' },
    }))).toContain(
      '<a:prstGeom prst="diamond"><a:avLst/></a:prstGeom><a:noFill/><a:ln/>',
    );
  });

  it('renders strict shape lines without changing the default shape skeleton', () => {
    expect(renderPresetShapeXml(7, normalizePresetShape('rect', {
      line: { kind: 'none' },
    }))).toContain(
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
      '<a:noFill/><a:ln><a:noFill/></a:ln>',
    );
    expect(renderPresetShapeXml(8, normalizePresetShape('ellipse', {
      fill: {
        kind: 'solid',
        color: { kind: 'srgb', value: 'EEEEEE' },
      },
      line: {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent2' },
        transparency: 25,
        width: 2.5,
        dash: 'dashDot',
      },
    }))).toContain(
      '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>' +
      '<a:solidFill><a:srgbClr val="EEEEEE"/></a:solidFill>' +
      '<a:ln w="31750"><a:solidFill><a:schemeClr val="accent2">' +
      '<a:alpha val="75000"/></a:schemeClr></a:solidFill>' +
      '<a:prstDash val="dashDot"/></a:ln>',
    );
    expect(renderPresetShapeXml(9, normalizePresetShape('diamond', {
      line: {
        kind: 'line',
        color: { kind: 'srgb', value: 'FF0000' },
        transparency: 0,
        width: 0,
      },
    }))).toContain(
      '<a:ln w="0"><a:solidFill><a:srgbClr val="FF0000">' +
      '<a:alpha val="100000"/></a:srgbClr></a:solidFill>' +
      '<a:prstDash val="solid"/></a:ln>',
    );
  });

  it('renders strict arrows after line state without synthesizing line defaults', () => {
    expect(renderPresetShapeXml(7, normalizePresetShape('line', {
      arrows: { begin: 'triangle', end: 'arrow' },
    }))).toContain(
      '<a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:noFill/>' +
      '<a:ln><a:headEnd type="triangle"/><a:tailEnd type="arrow"/></a:ln>',
    );
    expect(renderPresetShapeXml(8, normalizePresetShape('lineInv', {
      line: { kind: 'none' },
      arrows: { end: 'none' },
    }))).toContain(
      '<a:ln><a:noFill/><a:tailEnd type="none"/></a:ln>',
    );
    expect(renderPresetShapeXml(9, normalizePresetShape('line', {
      line: {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent2' },
        width: 2.5,
        dash: 'dashDot',
      },
      arrows: { begin: 'stealth', end: 'oval' },
    }))).toContain(
      '<a:ln w="31750"><a:solidFill><a:schemeClr val="accent2"/>' +
      '</a:solidFill><a:prstDash val="dashDot"/>' +
      '<a:headEnd type="stealth"/><a:tailEnd type="oval"/></a:ln>',
    );
    expect(renderPresetShapeXml(10, normalizePresetShape('line', {
      arrows: {},
    }))).toContain('<a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:noFill/><a:ln/>');
  });

  it('renders exact outer and inner shadows after line state in one effect list', () => {
    const outer = renderPresetShapeXml(7, normalizePresetShape('roundRect', {
      fill: { kind: 'solid', color: { kind: 'srgb', value: 'EEEEEE' } },
      line: {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent1' },
        width: 2,
      },
      arrows: { begin: 'triangle', end: 'arrow' },
      hyperlink: { url: 'https://example.com' },
      shadow: {
        kind: 'outer',
        color: { kind: 'srgb', value: '123ABC' },
        opacity: 0.42,
        blur: 7.25,
        angle: 123.4,
        distance: 5.5,
        rotateWithShape: true,
      },
    }), 'rId7');
    expect(outer).toContain(
      `r:id="rId7"`,
    );
    expect(outer).toContain(
      '<a:ln w="25400"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill>' +
      '<a:prstDash val="solid"/><a:headEnd type="triangle"/>' +
      '<a:tailEnd type="arrow"/></a:ln>' +
      '<a:effectLst><a:outerShdw sx="100000" sy="100000" kx="0" ky="0" ' +
      'algn="bl" rotWithShape="1" blurRad="92075" dist="69850" dir="7404000">' +
      '<a:srgbClr val="123ABC"><a:alpha val="42000"/></a:srgbClr>' +
      '</a:outerShdw></a:effectLst></p:spPr>',
    );
    expect(renderPresetShapeXml(8, normalizePresetShape('ellipse', {
      shadow: {
        kind: 'inner',
        color: { kind: 'scheme', value: 'accent2' },
        opacity: 0,
        blur: 0,
        angle: 0,
        distance: 0,
      },
    }))).toContain(
      '<a:ln/><a:effectLst><a:innerShdw blurRad="0" dist="0" dir="0">' +
      '<a:schemeClr val="accent2"><a:alpha val="0"/></a:schemeClr>' +
      '</a:innerShdw></a:effectLst></p:spPr>',
    );
  });

  it('renders ordered preset shape adjustments inside the existing geometry list', () => {
    const rendered = renderPresetShapeXml(7, normalizePresetShape('blockArc', {
      adjustments: [
        { name: 'adj1', value: 16_200_000 },
        { name: 'adj2', value: 0 },
        { name: 'adj3', value: 25_000 },
      ],
      fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent2' } },
      shadow: { kind: 'outer' },
    }));
    expect(rendered).toContain(
      '<a:prstGeom prst="blockArc"><a:avLst>' +
      '<a:gd name="adj1" fmla="val 16200000"/>' +
      '<a:gd name="adj2" fmla="val 0"/>' +
      '<a:gd name="adj3" fmla="val 25000"/></a:avLst></a:prstGeom>' +
      '<a:solidFill><a:schemeClr val="accent2"/></a:solidFill><a:ln/>' +
      '<a:effectLst>',
    );
    expect(renderPresetShapeXml(
      8,
      normalizePresetShape('roundRect', { adjustments: [] }),
    )).toContain('<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>');
  });

  it('creates every canonical geometry with one parseable default outer shadow', () => {
    for (const type of PRESET_SHAPE_TYPES) {
      const { xml } = parseShape(renderPresetShapeXml(
        7,
        normalizePresetShape(type, { shadow: { kind: 'outer' } }),
      ));
      expect(xml.elements('effectLst'), type).toHaveLength(1);
      expect(xml.elements('outerShdw'), type).toHaveLength(1);
      expect(xml.elements('innerShdw'), type).toHaveLength(0);
    }
  });

  it('round-trips every canonical token through parseable direct geometry', () => {
    for (const type of PRESET_SHAPE_TYPES) {
      const { xml, shape } = parseShape(renderPresetShapeXml(7, normalizePresetShape(type, {})));
      expect(readPresetShapeType(xml, shape), type).toBe(type);
      expect(xml.elements('prstGeom')).toHaveLength(1);
      expect(xml.elements('avLst')).toHaveLength(1);
      expect(xml.elements('noFill')).toHaveLength(1);
      expect(xml.elements('ln')).toHaveLength(1);
      expect(xml.elements('txBody')).toHaveLength(0);
      expect(xml.elements('extLst')).toHaveLength(0);
      expect(xml.source).not.toContain('r:id');
    }
  });

  it('reads alternate valid prefixes without changing source', () => {
    const source =
      `<q:sp xmlns:q="${PRESENTATION_NAMESPACE}" xmlns:d="${DRAWING_NAMESPACE}">` +
      '<q:spPr><d:prstGeom xmlns:prst="urn:unrelated" prst="star5">' +
      '<d:avLst/></d:prstGeom></q:spPr></q:sp>';
    const { xml, shape } = parseShape(source);
    expect(readPresetShapeType(xml, shape)).toBe('star5');
    expect(xml.changed).toBe(false);
    expect(xml.serialize()).toBe(source);
  });

  it('returns undefined for unsafe or ambiguous geometry without mutation', () => {
    const fixtures = [
      shapeFixture(undefined, { namespace: 'urn:wrong' }),
      shapeFixture(undefined, { rootName: 'p:pic' }),
      shapeFixture(undefined, { properties: '' }),
      shapeFixture(undefined, { properties: '<p:spPr/><p:spPr/>' }),
      shapeFixture(undefined, { properties: '<x:spPr xmlns:x="urn:wrong"/>' }),
      shapeFixture('', { properties: '<p:spPr/>' }),
      shapeFixture('<a:prstGeom prst="rect"/><a:prstGeom prst="ellipse"/>'),
      shapeFixture('<a:prstGeom prst="rect"/><a:custGeom/>'),
      shapeFixture('<x:prstGeom xmlns:x="urn:wrong" prst="rect"/>'),
      shapeFixture('<a:prstGeom/>'),
      shapeFixture('<a:prstGeom x:prst="rect" xmlns:x="urn:qualified"/>'),
      shapeFixture('<a:prstGeom prst="rect" x:prst="ellipse" xmlns:x="urn:qualified"/>'),
      shapeFixture('<a:prstGeom prst="folderCorner"/>'),
      shapeFixture('<a:prstGeom prst="custGeom"/>'),
      shapeFixture('<a:solidFill><a:prstGeom prst="rect"/></a:solidFill>'),
    ];
    for (const source of fixtures) {
      const { xml, shape } = parseShape(source);
      expect(readPresetShapeType(xml, shape), source).toBeUndefined();
      expect(xml.changed).toBe(false);
      expect(xml.serialize()).toBe(source);
    }
  });

  it('preserves same-value geometry exactly, including adjustments', () => {
    const source = shapeFixture(
      '<a:prstGeom prst="ellipse"><a:avLst><a:gd name="adj" fmla="val 1"/></a:avLst>' +
      '<x:keep xmlns:x="urn:test"/></a:prstGeom>',
    );
    const { xml, shape } = parseShape(source);
    expect(replacePresetShapeType(xml, shape, 'ellipse', '/ppt/slides/slide1.xml')).toBe(false);
    expect(xml.changed).toBe(false);
    expect(xml.serialize()).toBe(source);
  });

  it('whole-replaces changed geometry and preserves every unrelated byte', () => {
    const source = shapeFixture(
      '<a:prstGeom prst="rect"><a:avLst><a:gd name="adj" fmla="val 1"/></a:avLst>' +
      '<x:old xmlns:x="urn:test"/></a:prstGeom>',
      {
        properties:
          '<p:spPr data-keep="yes"><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm>' +
          '<a:prstGeom prst="rect"><a:avLst><a:gd name="adj" fmla="val 1"/></a:avLst>' +
          '<x:old xmlns:x="urn:test"/></a:prstGeom><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill>' +
          '<a:ln w="9"/><a:effectLst/><p:extLst><p:ext uri="urn:keep"/></p:extLst></p:spPr>',
      },
    );
    const { xml, shape } = parseShape(source);
    expect(replacePresetShapeType(xml, shape, 'star5', '/ppt/slides/slide1.xml')).toBe(true);
    const updated = xml.serialize();
    expect(updated).toContain('<a:prstGeom prst="star5"><a:avLst/></a:prstGeom>');
    expect(updated).not.toContain('name="adj"');
    expect(updated).not.toContain('<x:old');
    expect(updated).toContain('<p:cNvPr id="7" name="Keep"/>');
    expect(updated).toContain('<a:off x="1" y="2"/><a:ext cx="3" cy="4"/>');
    expect(updated).toContain('<a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill>');
    expect(updated).toContain('<a:ln w="9"/><a:effectLst/>');
    expect(updated).toContain('<p:extLst><p:ext uri="urn:keep"/></p:extLst>');
    expect(updated).toContain('<a:t>KEEP</a:t>');
  });

  it('converts supported custom geometry to a canonical preset geometry', () => {
    const source = renderCustomShapeXml(
      7,
      normalizeCustomShape(customGeometry, {
        name: 'Keep custom style',
        fill: { kind: 'solid', color: { kind: 'srgb', value: 'ABCDEF' } },
        line: { kind: 'line', color: { kind: 'scheme', value: 'accent1' } },
        shadow: { kind: 'outer' },
      }),
    );
    const { xml, shape } = parseShape(source);
    expect(readPresetShapeType(xml, shape)).toBeUndefined();
    expect(replacePresetShapeType(
      xml,
      shape,
      'ellipse',
      '/ppt/slides/slide1.xml',
    )).toBe(true);
    const updated = xml.serialize();
    const reparsed = parseShape(updated);
    expect(readPresetShapeType(reparsed.xml, reparsed.shape)).toBe('ellipse');
    expect(updated).toContain('<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>');
    expect(updated).not.toContain('<a:custGeom>');
    expect(updated).toContain('name="Keep custom style"');
    expect(updated).toContain('<a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill>');
    expect(updated).toContain('<a:ln w="12700"><a:solidFill><a:schemeClr val="accent1"/>');
    expect(updated).toContain('<a:effectLst>');
  });

  it('uses the existing DrawingML prefix when replacing geometry', () => {
    const source =
      `<q:sp xmlns:q="${PRESENTATION_NAMESPACE}" xmlns:d="${DRAWING_NAMESPACE}">` +
      '<q:spPr><d:prstGeom prst="rect"><d:avLst/></d:prstGeom></q:spPr></q:sp>';
    const { xml, shape } = parseShape(source);
    expect(replacePresetShapeType(xml, shape, 'ellipse', '/ppt/slides/slide1.xml')).toBe(true);
    expect(xml.serialize()).toContain('<d:prstGeom prst="ellipse"><d:avLst/></d:prstGeom>');
  });

  it('retains a DrawingML binding declared on the replaced geometry itself', () => {
    for (const geometry of [
      `<d:prstGeom xmlns:d="${DRAWING_NAMESPACE}" prst="rect"><d:avLst/></d:prstGeom>`,
      `<prstGeom xmlns="${DRAWING_NAMESPACE}" prst="rect"><avLst/></prstGeom>`,
    ]) {
      const source = `<p:sp xmlns:p="${PRESENTATION_NAMESPACE}"><p:spPr>` +
        `${geometry}</p:spPr></p:sp>`;
      const { xml, shape } = parseShape(source);
      expect(replacePresetShapeType(xml, shape, 'ellipse', '/ppt/slides/slide1.xml')).toBe(true);
      const reparsed = parseShape(xml.serialize());
      expect(readPresetShapeType(reparsed.xml, reparsed.shape)).toBe('ellipse');
    }
  });

  it('rejects unsafe replacement before patching', () => {
    for (const source of [
      shapeFixture('', { properties: '<p:spPr/>' }),
      shapeFixture('<a:prstGeom prst="folderCorner"/>'),
      shapeFixture('<a:prstGeom prst="rect"/><a:prstGeom prst="ellipse"/>'),
      shapeFixture('<a:prstGeom prst="rect"/><a:custGeom/>'),
      shapeFixture(
        '<a:custGeom><a:avLst/><a:gdLst><a:gd name="x" fmla="unknown 1"/></a:gdLst>' +
        '<a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/>' +
        '<a:pathLst><a:path w="1" h="1"/></a:pathLst></a:custGeom>',
      ),
    ]) {
      const { xml, shape } = parseShape(source);
      expect(() => replacePresetShapeType(
        xml,
        shape,
        'star5',
        '/ppt/slides/slide1.xml',
      )).toThrow(ModelParseError);
      expect(xml.changed).toBe(false);
      expect(xml.serialize()).toBe(source);
    }
  });
});
