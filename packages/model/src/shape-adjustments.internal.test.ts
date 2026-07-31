import { LosslessXmlDocument } from '@pptx/lossless-xml';
import { describe, expect, it } from 'vitest';
import { ModelParseError } from './errors.js';
import type { ShapeAdjustment } from './preset-shape.js';
import {
  normalizeShapeAdjustments,
  readShapeAdjustments,
  renderShapeAdjustmentList,
  replaceShapeAdjustments,
  shapeAdjustmentsEqual,
} from './shape-adjustments.internal.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const PART_URI = '/ppt/slides/slide1.xml';

function fixture(
  geometry = '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>',
  options: {
    readonly rootName?: string;
    readonly presentationNamespace?: string;
    readonly properties?: string;
  } = {},
): string {
  const rootName = options.rootName ?? 'p:sp';
  const properties = options.properties ??
    `<p:spPr keep="PROPERTIES"><a:xfrm/>${geometry}` +
    '<a:noFill/><a:ln/><a:effectLst/><p:extLst><p:ext uri="urn:keep"/>' +
    '</p:extLst></p:spPr>';
  return `<${rootName} xmlns:p="${options.presentationNamespace ?? PRESENTATION_NAMESPACE}" ` +
    `xmlns:a="${DRAWING_NAMESPACE}">` +
    '<p:nvSpPr><p:cNvPr id="7" name="Keep"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
    `${properties}<p:txBody><a:bodyPr/><a:p><a:r><a:t>KEEP</a:t></a:r></a:p></p:txBody>` +
    `</${rootName}>`;
}

function parse(source: string) {
  const xml = LosslessXmlDocument.parse(source);
  const shape = xml.roots[0];
  if (!shape) throw new Error('Fixture has no shape');
  return { xml, shape };
}

describe('shape adjustment normalization', () => {
  it('normalizes ordered values into detached deeply frozen snapshots', () => {
    const nullPrototype = Object.assign(Object.create(null), {
      name: 'adj2',
      value: 0,
    }) as ShapeAdjustment;
    const input = [
      { name: 'adj&<"', value: 16_200_000 },
      nullPrototype,
      { name: 'minimum', value: Number.MIN_SAFE_INTEGER },
      { name: 'maximum', value: Number.MAX_SAFE_INTEGER },
      { name: 'negative', value: -7 },
    ];

    const normalized = normalizeShapeAdjustments(input, 'Shape adjustments');
    (input[0] as { name: string; value: number }).name = 'changed';
    (input[0] as { name: string; value: number }).value = 1;
    (nullPrototype as { name: string; value: number }).value = 2;

    expect(normalized).toEqual([
      { name: 'adj&<"', value: 16_200_000 },
      { name: 'adj2', value: 0 },
      { name: 'minimum', value: Number.MIN_SAFE_INTEGER },
      { name: 'maximum', value: Number.MAX_SAFE_INTEGER },
      { name: 'negative', value: -7 },
    ]);
    expect(Object.isFrozen(normalized)).toBe(true);
    expect(normalized.every(Object.isFrozen)).toBe(true);
    expect(normalizeShapeAdjustments([], 'Shape adjustments')).toEqual([]);
    expect(Object.isFrozen(normalizeShapeAdjustments([], 'Shape adjustments'))).toBe(true);
  });

  it('rejects non-ordinary, sparse, accessor, symbol, and extended arrays', () => {
    class AdjustmentArray<T> extends Array<T> {}
    const sparse = new Array(2);
    sparse[1] = { name: 'adj', value: 1 };
    const extended = [{ name: 'adj', value: 1 }];
    Object.defineProperty(extended, 'extra', { value: true });
    const symbol = [{ name: 'adj', value: 1 }];
    Object.defineProperty(symbol, Symbol('extra'), { value: true });
    let reads = 0;
    const accessor = [{ name: 'adj', value: 1 }];
    Object.defineProperty(accessor, '0', {
      configurable: true,
      enumerable: true,
      get() {
        reads += 1;
        return { name: 'adj', value: 1 };
      },
    });

    for (const value of [
      undefined,
      null,
      true,
      1,
      'x',
      {},
      new Set(),
      new AdjustmentArray({ name: 'adj', value: 1 }),
      sparse,
      extended,
      symbol,
      accessor,
    ]) {
      expect(() => normalizeShapeAdjustments(value, 'Shape adjustments')).toThrow(TypeError);
    }
    expect(reads).toBe(0);
  });

  it('rejects malformed entries without invoking accessors', () => {
    class Entry {
      name = 'adj';
      value = 1;
    }
    let reads = 0;
    const accessor = Object.defineProperty({ name: 'adj' }, 'value', {
      enumerable: true,
      get() {
        reads += 1;
        return 1;
      },
    });
    const symbol = { name: 'adj', value: 1 };
    Object.defineProperty(symbol, Symbol('extra'), { value: true });
    const inherited = Object.create({ name: 'adj', value: 1 });

    const values = [
      [null],
      [[]],
      [new Date()],
      [new Entry()],
      [inherited],
      [accessor],
      [symbol],
      [{}],
      [{ name: 'adj' }],
      [{ value: 1 }],
      [{ name: 'adj', value: 1, extra: true }],
      [{ name: '', value: 1 }],
      [{ name: 7, value: 1 }],
      [{ name: 'bad\u0000name', value: 1 }],
      [{ name: 'bad\uD800name', value: 1 }],
      [{ name: 'bad\uFFFEname', value: 1 }],
      [{ name: 'adj', value: '1' }],
      [{ name: 'adj', value: 1.5 }],
      [{ name: 'adj', value: Number.NaN }],
      [{ name: 'adj', value: Number.POSITIVE_INFINITY }],
      [{ name: 'adj', value: Number.MAX_SAFE_INTEGER + 1 }],
      [{ name: 'adj', value: 1 }, { name: 'adj', value: 2 }],
    ];
    for (const value of values) {
      expect(() => normalizeShapeAdjustments(value, 'Shape adjustments')).toThrow();
    }
    expect(reads).toBe(0);
  });
});

describe('shape adjustment rendering and equality', () => {
  it('renders deterministic same-prefix compact XML', () => {
    expect(renderShapeAdjustmentList([], 'd:')).toBe('<d:avLst/>');
    expect(renderShapeAdjustmentList(normalizeShapeAdjustments([
      { name: 'adj&<"', value: 16_200_000 },
      { name: 'adj2', value: 0 },
      { name: 'adj3', value: -7 },
    ], 'Shape adjustments'), 'a:')).toBe(
      '<a:avLst><a:gd name="adj&amp;&lt;&quot;" fmla="val 16200000"/>' +
      '<a:gd name="adj2" fmla="val 0"/><a:gd name="adj3" fmla="val -7"/>' +
      '</a:avLst>',
    );
  });

  it('compares complete ordered values', () => {
    const value = normalizeShapeAdjustments([
      { name: 'adj1', value: 1 },
      { name: 'adj2', value: 2 },
    ], 'Shape adjustments');
    expect(shapeAdjustmentsEqual(value, [
      { name: 'adj1', value: 1 },
      { name: 'adj2', value: 2 },
    ])).toBe(true);
    expect(shapeAdjustmentsEqual(value, [...value].reverse())).toBe(false);
    expect(shapeAdjustmentsEqual(value, [{ name: 'adj1', value: 1 }])).toBe(false);
    expect(shapeAdjustmentsEqual(value, [
      { name: 'renamed', value: 1 },
      { name: 'adj2', value: 2 },
    ])).toBe(false);
    expect(shapeAdjustmentsEqual(value, [
      { name: 'adj1', value: 7 },
      { name: 'adj2', value: 2 },
    ])).toBe(false);
    expect(shapeAdjustmentsEqual(undefined, undefined)).toBe(true);
    expect(shapeAdjustmentsEqual(value, undefined)).toBe(false);
  });
});

describe('shape adjustment reader', () => {
  it('reads canonical and alternate-prefix values as detached frozen snapshots', () => {
    const sources = [
      fixture(),
      fixture(
        '<a:prstGeom prst="blockArc"><a:avLst>' +
        '<a:gd name="adj1" fmla="val 16200000"/>' +
        '<a:gd name="adj2" fmla="val\t+0"/>' +
        '<a:gd name="adj3" fmla="val&#xA;-0"/>' +
        '</a:avLst></a:prstGeom>',
      ),
      `<q:sp xmlns:q="${PRESENTATION_NAMESPACE}" xmlns:d="${DRAWING_NAMESPACE}">` +
        '<q:spPr><d:prstGeom prst="pie"><d:avLst xmlns:k="urn:keep">' +
        '<d:gd name="adj1" fmla="val&#xD;5400000"/>' +
        '<d:gd name="adj2" fmla="val 0"/></d:avLst></d:prstGeom></q:spPr></q:sp>',
    ];
    const expected = [
      [],
      [
        { name: 'adj1', value: 16_200_000 },
        { name: 'adj2', value: 0 },
        { name: 'adj3', value: 0 },
      ],
      [
        { name: 'adj1', value: 5_400_000 },
        { name: 'adj2', value: 0 },
      ],
    ];

    for (let index = 0; index < sources.length; index += 1) {
      const source = sources[index]!;
      const { xml, shape } = parse(source);
      const first = readShapeAdjustments(xml, shape);
      const second = readShapeAdjustments(xml, shape);
      expect(first, source).toEqual(expected[index]);
      expect(second, source).toEqual(expected[index]);
      expect(first).not.toBe(second);
      expect(Object.isFrozen(first)).toBe(true);
      expect(first?.every(Object.isFrozen)).toBe(true);
      expect(xml.changed).toBe(false);
      expect(xml.serialize()).toBe(source);
    }
  });

  it('returns undefined for unsafe owner, list, guide, attribute, and formula state', () => {
    const geometries = [
      '',
      '<a:prstGeom prst="rect"/>',
      '<a:prstGeom prst="folderCorner"><a:avLst/></a:prstGeom>',
      '<a:prstGeom prst="rect"><a:avLst/><a:avLst/></a:prstGeom>',
      '<a:prstGeom prst="rect"><x:avLst xmlns:x="urn:wrong"/></a:prstGeom>',
      '<a:prstGeom prst="rect"><a:avLst custom="x"/></a:prstGeom>',
      '<a:prstGeom prst="rect"><a:avLst>TEXT</a:avLst></a:prstGeom>',
      '<a:prstGeom prst="rect"><a:avLst>&#xA0;</a:avLst></a:prstGeom>',
      '<a:prstGeom prst="rect"><a:avLst><a:ext/></a:avLst></a:prstGeom>',
      '<a:prstGeom prst="rect"><a:avLst><x:gd xmlns:x="urn:wrong" name="adj" fmla="val 1"/></a:avLst></a:prstGeom>',
      '<a:prstGeom prst="rect"><a:avLst><a:gd/></a:avLst></a:prstGeom>',
      '<a:prstGeom prst="rect"><a:avLst><a:gd name="adj"/></a:avLst></a:prstGeom>',
      '<a:prstGeom prst="rect"><a:avLst><a:gd fmla="val 1"/></a:avLst></a:prstGeom>',
      '<a:prstGeom prst="rect"><a:avLst><a:gd x:name="adj" xmlns:x="urn:q" fmla="val 1"/></a:avLst></a:prstGeom>',
      '<a:prstGeom prst="rect"><a:avLst><a:gd name="adj" x:fmla="val 1" xmlns:x="urn:q"/></a:avLst></a:prstGeom>',
      '<a:prstGeom prst="rect"><a:avLst><a:gd name="adj" fmla="val 1" custom="x"/></a:avLst></a:prstGeom>',
      '<a:prstGeom prst="rect"><a:avLst><a:gd name="adj" fmla="val 1"><a:ext/></a:gd></a:avLst></a:prstGeom>',
      '<a:prstGeom prst="rect"><a:avLst><a:gd name="adj" fmla="val 1">TEXT</a:gd></a:avLst></a:prstGeom>',
      '<a:prstGeom prst="rect"><a:avLst><a:gd name="adj" fmla="val 1">&#xA0;</a:gd></a:avLst></a:prstGeom>',
      '<a:prstGeom prst="rect"><a:avLst><a:gd name="adj" fmla="val1"/></a:avLst></a:prstGeom>',
      '<a:prstGeom prst="rect"><a:avLst><a:gd name="adj" fmla=" val 1"/></a:avLst></a:prstGeom>',
      '<a:prstGeom prst="rect"><a:avLst><a:gd name="adj" fmla="val 1 "/></a:avLst></a:prstGeom>',
      '<a:prstGeom prst="rect"><a:avLst><a:gd name="adj" fmla="val 1 2"/></a:avLst></a:prstGeom>',
      '<a:prstGeom prst="rect"><a:avLst><a:gd name="adj" fmla="val 1.5"/></a:avLst></a:prstGeom>',
      '<a:prstGeom prst="rect"><a:avLst><a:gd name="adj" fmla="val 1e3"/></a:avLst></a:prstGeom>',
      '<a:prstGeom prst="rect"><a:avLst><a:gd name="adj" fmla="val 9007199254740992"/></a:avLst></a:prstGeom>',
      '<a:prstGeom prst="rect"><a:avLst><a:gd name="adj" fmla="*/ 1 2 3"/></a:avLst></a:prstGeom>',
      '<a:prstGeom prst="rect"><a:avLst><a:gd name="adj" fmla="pin 0 x 1"/></a:avLst></a:prstGeom>',
      '<a:prstGeom prst="rect"><a:avLst><a:gd name="adj" fmla="val 1"/><a:gd name="adj" fmla="val 2"/></a:avLst></a:prstGeom>',
    ];
    const sources = [
      fixture(undefined, { rootName: 'p:pic' }),
      fixture(undefined, { presentationNamespace: 'urn:wrong' }),
      fixture(undefined, { properties: '<p:spPr/><p:spPr/>' }),
      fixture(undefined, { properties: '<x:spPr xmlns:x="urn:wrong"/>' }),
      fixture('<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
        '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>'),
      fixture('<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:custGeom/>'),
      ...geometries.map((geometry) => fixture(geometry)),
    ];

    for (const source of sources) {
      const { xml, shape } = parse(source);
      expect(readShapeAdjustments(xml, shape), source).toBeUndefined();
      expect(xml.changed, source).toBe(false);
      expect(xml.serialize(), source).toBe(source);
    }
  });
});

describe('shape adjustment replacement', () => {
  it('preserves exact bytes for semantically equal ordered values', () => {
    const source = fixture(
      '<a:prstGeom prst="pie"><a:avLst>' +
      '<a:gd name="adj1" fmla="val\t+5400000"/>' +
      '<a:gd name="adj2" fmla="val -0"/></a:avLst>' +
      '<x:keep xmlns:x="urn:test"/></a:prstGeom>',
    );
    const { xml, shape } = parse(source);
    const value = normalizeShapeAdjustments([
      { name: 'adj1', value: 5_400_000 },
      { name: 'adj2', value: 0 },
    ], 'Shape adjustments');
    expect(replaceShapeAdjustments(xml, shape, value, PART_URI)).toBe(false);
    expect(xml.changed).toBe(false);
    expect(xml.serialize()).toBe(source);
  });

  it('replaces only the adjustment list and clears with the existing prefix', () => {
    const source =
      `<q:sp xmlns:q="${PRESENTATION_NAMESPACE}" xmlns:d="${DRAWING_NAMESPACE}">` +
      '<q:nvSpPr><q:cNvPr id="7" name="Keep"/></q:nvSpPr>' +
      '<q:spPr keep="PROPERTIES"><d:xfrm/><d:prstGeom prst="blockArc" custom="KEEP">' +
      '<d:avLst><d:gd name="adj1" fmla="val 1"/></d:avLst>' +
      '<x:keep xmlns:x="urn:test">KEEP</x:keep></d:prstGeom>' +
      '<d:solidFill><d:srgbClr val="ABCDEF"/></d:solidFill><d:ln w="9"/>' +
      '<d:effectLst/><q:extLst><q:ext uri="urn:keep"/></q:extLst></q:spPr>' +
      '<q:txBody><d:bodyPr/><d:p><d:r><d:t>KEEP</d:t></d:r></d:p></q:txBody></q:sp>';
    const { xml, shape } = parse(source);
    const replacement = normalizeShapeAdjustments([
      { name: 'adj1', value: 16_200_000 },
      { name: 'adj2', value: 0 },
      { name: 'adj3', value: 25_000 },
    ], 'Shape adjustments');
    expect(replaceShapeAdjustments(xml, shape, replacement, PART_URI)).toBe(true);
    const updated = xml.serialize();
    expect(updated).toContain(
      '<d:prstGeom prst="blockArc" custom="KEEP"><d:avLst>' +
      '<d:gd name="adj1" fmla="val 16200000"/>' +
      '<d:gd name="adj2" fmla="val 0"/>' +
      '<d:gd name="adj3" fmla="val 25000"/></d:avLst>' +
      '<x:keep xmlns:x="urn:test">KEEP</x:keep></d:prstGeom>',
    );
    expect(updated).toContain('<q:cNvPr id="7" name="Keep"/>');
    expect(updated).toContain('<d:solidFill><d:srgbClr val="ABCDEF"/></d:solidFill>');
    expect(updated).toContain('<d:ln w="9"/><d:effectLst/>');
    expect(updated).toContain('<q:extLst><q:ext uri="urn:keep"/></q:extLst>');
    expect(updated).toContain('<d:t>KEEP</d:t>');

    const reparsed = parse(updated);
    expect(replaceShapeAdjustments(
      reparsed.xml,
      reparsed.shape,
      normalizeShapeAdjustments([], 'Shape adjustments'),
      PART_URI,
    )).toBe(true);
    expect(reparsed.xml.serialize()).toContain(
      '<d:prstGeom prst="blockArc" custom="KEEP"><d:avLst/>' +
      '<x:keep xmlns:x="urn:test">KEEP</x:keep></d:prstGeom>',
    );
  });

  it('retains a DrawingML namespace binding declared on the list', () => {
    const source =
      `<p:sp xmlns:p="${PRESENTATION_NAMESPACE}" xmlns:a="${DRAWING_NAMESPACE}">` +
      '<p:spPr><a:prstGeom prst="rect">' +
      `<d:avLst xmlns:d="${DRAWING_NAMESPACE}">` +
      '<d:gd name="adj" fmla="val 1"/></d:avLst></a:prstGeom>' +
      '</p:spPr></p:sp>';
    const { xml, shape } = parse(source);
    expect(replaceShapeAdjustments(
      xml,
      shape,
      normalizeShapeAdjustments([{ name: 'adj', value: 2 }], 'Shape adjustments'),
      PART_URI,
    )).toBe(true);
    const reparsed = parse(xml.serialize());
    expect(readShapeAdjustments(reparsed.xml, reparsed.shape)).toEqual([
      { name: 'adj', value: 2 },
    ]);
  });

  it('rejects unsupported source state before patching', () => {
    const sources = [
      fixture('<a:prstGeom prst="rect"/>'),
      fixture('<a:prstGeom prst="rect"><a:avLst>' +
        '<a:gd name="adj" fmla="*/ 1 2 3"/></a:avLst></a:prstGeom>'),
      fixture('<a:prstGeom prst="rect"><a:avLst/><a:avLst/></a:prstGeom>'),
      fixture('<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:custGeom/>'),
      fixture('<a:prstGeom prst="rect"><a:avLst><a:gd name="adj" fmla="val 1"/>' +
        '<a:gd name="adj" fmla="val 2"/></a:avLst></a:prstGeom>'),
    ];
    const value = normalizeShapeAdjustments([{ name: 'adj', value: 7 }], 'Shape adjustments');
    for (const source of sources) {
      const { xml, shape } = parse(source);
      expect(() => replaceShapeAdjustments(xml, shape, value, PART_URI))
        .toThrow(ModelParseError);
      expect(() => replaceShapeAdjustments(xml, shape, value, PART_URI))
        .toThrow(PART_URI);
      expect(xml.changed).toBe(false);
      expect(xml.serialize()).toBe(source);
    }
  });
});
