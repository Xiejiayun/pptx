import { describe, expect, it } from 'vitest';
import { LosslessXmlDocument } from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
import type {
  AddCustomShapeOptions,
  CustomGeometry,
  CustomGeometryCommand,
} from './custom-geometry.js';
import { inches } from './units.js';
import {
  customGeometryEqual,
  normalizeCustomGeometry,
  readCustomGeometry,
  renderCustomGeometry,
  replaceCustomGeometry,
} from './custom-geometry.internal.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';

const allCommands: readonly CustomGeometryCommand[] = [
  { kind: 'moveTo', point: { x: 0, y: 0 } },
  { kind: 'lineTo', point: { x: 914_400, y: 0 } },
  {
    kind: 'quadraticBezierTo',
    control: { x: 1_371_600, y: 0 },
    end: { x: 1_828_800, y: 914_400 },
  },
  {
    kind: 'cubicBezierTo',
    control1: { x: 2_057_400, y: 914_400 },
    control2: { x: 2_514_600, y: 1_828_800 },
    end: { x: 2_743_200, y: 1_828_800 },
  },
  {
    kind: 'arcTo',
    widthRadius: 914_400,
    heightRadius: 457_200,
    startAngle: 1_800_000,
    sweepAngle: 7_200_000,
  },
  { kind: 'moveTo', point: { x: -1, y: -2 } },
  { kind: 'lineTo', point: { x: 0, y: 0 } },
  { kind: 'close' },
];

const geometry: CustomGeometry = {
  paths: [
    {
      width: 3_657_600,
      height: 2_743_200,
      fill: 'norm',
      stroke: true,
      extrusionOk: false,
      commands: allCommands,
    },
    { width: 914_400, height: 914_400, commands: [] },
  ],
};

const publicOptions: AddCustomShapeOptions = { name: 'Custom', x: inches(1) };
void publicOptions;
// @ts-expect-error preset-only adjustments are not custom-shape options
const invalidOptions: AddCustomShapeOptions = { adjustments: [] };
void invalidOptions;
// @ts-expect-error custom geometry coordinates are numeric direct values
const invalidCommand: CustomGeometryCommand = { kind: 'moveTo', point: { x: '1', y: 2 } };
void invalidCommand;
// @ts-expect-error custom geometry commands must use a supported kind
const unknownCommand: CustomGeometryCommand = { kind: 'unknown' };
void unknownCommand;
// @ts-expect-error move commands require a point
const missingPoint: CustomGeometryCommand = { kind: 'moveTo' };
void missingPoint;
const arcEndpoint: CustomGeometryCommand = {
  kind: 'arcTo',
  widthRadius: 1,
  heightRadius: 1,
  startAngle: 0,
  sweepAngle: 0,
  // @ts-expect-error arc commands do not accept an endpoint
  end: { x: 1, y: 1 },
};
void arcEndpoint;

function parseShape(source: string) {
  const xml = LosslessXmlDocument.parse(source);
  const shape = xml.roots[0];
  if (!shape) throw new Error('Fixture has no shape root');
  return { xml, shape };
}

function fixture(
  customGeometry: string,
  options: { readonly drawingPrefix?: string; readonly drawingNamespace?: string } = {},
): string {
  const drawingPrefix = options.drawingPrefix ?? 'a';
  const qualified = drawingPrefix === '' ? '' : `${drawingPrefix}:`;
  const declaration = drawingPrefix === ''
    ? `xmlns="${options.drawingNamespace ?? DRAWING_NAMESPACE}"`
    : `xmlns:${drawingPrefix}="${options.drawingNamespace ?? DRAWING_NAMESPACE}"`;
  return `<p:sp xmlns:p="${PRESENTATION_NAMESPACE}" ${declaration}>` +
    '<p:nvSpPr><p:cNvPr id="2" name="Keep"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
    `<p:spPr><${qualified}xfrm/><${qualified}solidFill/>${customGeometry}` +
    `<${qualified}ln/><${qualified}effectLst/><${qualified}extLst/></p:spPr>` +
    `<p:txBody><${qualified}bodyPr/><${qualified}p/></p:txBody></p:sp>`;
}

function canonical(prefix = 'a:'): string {
  return `<${prefix}custGeom><${prefix}avLst/><${prefix}gdLst/><${prefix}ahLst/>` +
    `<${prefix}cxnLst/><${prefix}rect l="l" t="t" r="r" b="b"/>` +
    `<${prefix}pathLst>` +
    `<${prefix}path w="3657600" h="2743200" fill="norm" stroke="1" extrusionOk="0">` +
    `<${prefix}moveTo><${prefix}pt x="0" y="0"/></${prefix}moveTo>` +
    `<${prefix}lnTo><${prefix}pt x="914400" y="0"/></${prefix}lnTo>` +
    `<${prefix}quadBezTo><${prefix}pt x="1371600" y="0"/>` +
    `<${prefix}pt x="1828800" y="914400"/></${prefix}quadBezTo>` +
    `<${prefix}cubicBezTo><${prefix}pt x="2057400" y="914400"/>` +
    `<${prefix}pt x="2514600" y="1828800"/>` +
    `<${prefix}pt x="2743200" y="1828800"/></${prefix}cubicBezTo>` +
    `<${prefix}arcTo wR="914400" hR="457200" stAng="1800000" swAng="7200000"/>` +
    `<${prefix}moveTo><${prefix}pt x="-1" y="-2"/></${prefix}moveTo>` +
    `<${prefix}lnTo><${prefix}pt x="0" y="0"/></${prefix}lnTo>` +
    `<${prefix}close/></${prefix}path>` +
    `<${prefix}path w="914400" h="914400"></${prefix}path>` +
    `</${prefix}pathLst></${prefix}custGeom>`;
}

describe('normalizeCustomGeometry', () => {
  it('copies, normalizes negative zero, and deeply freezes all path state', () => {
    const mutable = structuredClone(geometry) as {
      paths: Array<{
        width: number;
        height: number;
        fill?: 'norm';
        stroke?: boolean;
        extrusionOk?: boolean;
        commands: Array<CustomGeometryCommand>;
      }>;
    };
    const first = mutable.paths[0]!;
    first.commands[0] = { kind: 'moveTo', point: { x: -0, y: 0 } };
    const normalized = normalizeCustomGeometry(mutable, 'Custom geometry');
    expect(normalized).toEqual(geometry);
    expect(Object.isFrozen(normalized)).toBe(true);
    expect(Object.isFrozen(normalized.paths)).toBe(true);
    expect(Object.isFrozen(normalized.paths[0])).toBe(true);
    expect(Object.isFrozen(normalized.paths[0]?.commands)).toBe(true);
    expect(normalized.paths[0]?.commands.every(Object.isFrozen)).toBe(true);
    const firstCommand = normalized.paths[0]?.commands[0];
    expect(firstCommand?.kind).toBe('moveTo');
    if (firstCommand?.kind === 'moveTo') {
      expect(Object.isFrozen(firstCommand.point)).toBe(true);
      expect(Object.is(firstCommand.point.x, -0)).toBe(false);
    }

    first.width = 1;
    first.commands.splice(0);
    mutable.paths.splice(0);
    expect(normalized).toEqual(geometry);
  });

  it('accepts null-prototype values, later subpaths, empty paths, and signed values', () => {
    const point = Object.assign(Object.create(null), { x: -7, y: 8 });
    const move = Object.assign(Object.create(null), { kind: 'moveTo', point });
    const arc = Object.assign(Object.create(null), {
      kind: 'arcTo',
      widthRadius: 1,
      heightRadius: 2,
      startAngle: 0,
      sweepAngle: -3,
    });
    const path = Object.assign(Object.create(null), {
      width: 10,
      height: 20,
      commands: [move, arc, { kind: 'moveTo', point: { x: 1, y: 2 } }],
    });
    const root = Object.assign(Object.create(null), { paths: [path, {
      width: 1,
      height: 1,
      commands: [],
    }] });
    expect(normalizeCustomGeometry(root, 'Custom geometry')).toEqual(root);
  });

  it('rejects unsafe containers, arrays, properties, and accessors without invoking them', () => {
    class Geometry {}
    for (const value of [undefined, null, false, 1, 'x', [], new Date(), new Geometry()]) {
      expect(() => normalizeCustomGeometry(value, 'Custom geometry')).toThrow(TypeError);
    }
    expect(() => normalizeCustomGeometry(Object.create({ paths: [] }), 'Custom geometry'))
      .toThrow(/ordinary object/);
    expect(() => normalizeCustomGeometry({ paths: [], extra: true }, 'Custom geometry'))
      .toThrow(/unsupported property extra/);
    expect(() => normalizeCustomGeometry({ paths: [], [Symbol('unsafe')]: true }, 'Custom geometry'))
      .toThrow(/unsupported property/);
    const protoKey = Object.assign(Object.create(null), { paths: [geometry.paths[0]] });
    Object.defineProperty(protoKey, '__proto__', { value: true, enumerable: true });
    expect(() => normalizeCustomGeometry(protoKey, 'Custom geometry'))
      .toThrow(/unsupported property __proto__/);

    let calls = 0;
    const accessor = Object.defineProperty({}, 'paths', {
      enumerable: true,
      get() {
        calls += 1;
        return [];
      },
    });
    expect(() => normalizeCustomGeometry(accessor, 'Custom geometry')).toThrow(/data property/);
    const pointAccessor = Object.defineProperties({}, {
      x: {
        enumerable: true,
        get() {
          calls += 1;
          return 0;
        },
      },
      y: { enumerable: true, value: 0 },
    });
    expect(() => normalizeCustomGeometry({
      paths: [{
        width: 1,
        height: 1,
        commands: [{ kind: 'moveTo', point: pointAccessor }],
      }],
    }, 'Custom geometry')).toThrow(/data property/);
    const commands: unknown[] = [];
    Object.defineProperty(commands, '0', {
      enumerable: true,
      get() {
        calls += 1;
        return { kind: 'moveTo', point: { x: 0, y: 0 } };
      },
    });
    expect(() => normalizeCustomGeometry({
      paths: [{ width: 1, height: 1, commands }],
    }, 'Custom geometry')).toThrow(/data property/);
    expect(calls).toBe(0);
  });

  it('rejects sparse/exotic arrays and invalid path values', () => {
    const sparse = Array(1);
    expect(() => normalizeCustomGeometry({ paths: sparse }, 'Custom geometry'))
      .toThrow(/dense array/);
    const extra = [geometry.paths[0]] as unknown[] & { unsafe?: boolean };
    extra.unsafe = true;
    expect(() => normalizeCustomGeometry({ paths: extra }, 'Custom geometry'))
      .toThrow(/extra properties/);
    class Paths<T> extends Array<T> {}
    expect(() => normalizeCustomGeometry({ paths: new Paths() }, 'Custom geometry'))
      .toThrow(/ordinary array/);

    for (const candidate of [
      { paths: [] },
      { paths: [{ width: 0, height: 1, commands: [] }] },
      { paths: [{ width: 1, height: -1, commands: [] }] },
      { paths: [{ width: 1.5, height: 1, commands: [] }] },
      { paths: [{ width: Number.MAX_SAFE_INTEGER + 1, height: 1, commands: [] }] },
      { paths: [{ width: 1, height: 1, fill: 'invalid', commands: [] }] },
      { paths: [{ width: 1, height: 1, stroke: 1, commands: [] }] },
      { paths: [{ width: 1, height: 1, extrusionOk: 'false', commands: [] }] },
      { paths: [{ width: 1, height: 1, commands: [], extra: true }] },
    ]) expect(() => normalizeCustomGeometry(candidate, 'Custom geometry')).toThrow();
  });

  it('rejects invalid command branches, numbers, and sequence', () => {
    const wrap = (command: unknown) => ({
      paths: [{ width: 1, height: 1, commands: [command] }],
    });
    for (const command of [
      { kind: 'close' },
      { kind: 'lineTo', point: { x: 0, y: 0 } },
      { kind: 'arcTo', widthRadius: 1, heightRadius: 1, startAngle: 0, sweepAngle: 0 },
      { kind: 'unknown' },
      { kind: 'moveTo' },
      { kind: 'moveTo', point: { x: 0, y: 0 }, extra: true },
      { kind: 'moveTo', point: { x: NaN, y: 0 } },
      { kind: 'moveTo', point: { x: 1.5, y: 0 } },
      { kind: 'quadraticBezierTo', control: { x: 0, y: 0 } },
      { kind: 'cubicBezierTo', control1: { x: 0, y: 0 }, control2: { x: 0, y: 0 } },
    ]) expect(() => normalizeCustomGeometry(wrap(command), 'Custom geometry')).toThrow();

    const leadingMove = { kind: 'moveTo', point: { x: 0, y: 0 } };
    for (const command of [
      { kind: 'arcTo', widthRadius: 0, heightRadius: 1, startAngle: 0, sweepAngle: 0 },
      { kind: 'arcTo', widthRadius: 1, heightRadius: -1, startAngle: 0, sweepAngle: 0 },
      { kind: 'arcTo', widthRadius: 1, heightRadius: 1, startAngle: 0.5, sweepAngle: 0 },
      { kind: 'arcTo', widthRadius: 1, heightRadius: 1, startAngle: 0, sweepAngle: Infinity },
    ]) {
      expect(() => normalizeCustomGeometry({
        paths: [{ width: 1, height: 1, commands: [leadingMove, command] }],
      }, 'Custom geometry')).toThrow();
    }
  });
});

describe('custom geometry OOXML codec', () => {
  it('renders every command, multiple paths, and direct path flags canonically', () => {
    const normalized = normalizeCustomGeometry(geometry, 'Custom geometry');
    expect(renderCustomGeometry(normalized, 'a:')).toBe(canonical());
    expect(renderCustomGeometry(normalized, '')).toBe(canonical(''));
  });

  it('reads canonical, alternate-prefix, absent-empty-list, and boolean lexical state', () => {
    for (const source of [
      fixture(canonical()),
      fixture(canonical('d:'), { drawingPrefix: 'd' }),
      fixture(canonical()
        .replace('<a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>', '')
        .replace('<a:rect l="l" t="t" r="r" b="b"/>', '')),
      fixture(canonical()
        .replace('stroke="1" extrusionOk="0"', 'stroke="true" extrusionOk="off"')),
    ]) {
      const { xml, shape } = parseShape(source);
      expect(readCustomGeometry(xml, shape)).toEqual(geometry);
    }
  });

  it('distinguishes ordered paths, commands, values, flags, and optional absence', () => {
    const normalized = normalizeCustomGeometry(geometry, 'Custom geometry');
    expect(customGeometryEqual(normalized, normalized)).toBe(true);
    expect(customGeometryEqual(undefined, undefined)).toBe(true);
    expect(customGeometryEqual(normalized, undefined)).toBe(false);
    const { fill: _fill, ...pathWithoutFill } = geometry.paths[0]!;
    const variants = [
      { ...geometry, paths: [...geometry.paths].reverse() },
      { paths: [{ ...geometry.paths[0]!, width: 1 }] },
      { paths: [pathWithoutFill] },
      { paths: [{ ...geometry.paths[0]!, stroke: false }] },
      { paths: [{ ...geometry.paths[0]!, commands: [
        allCommands[0]!, allCommands[2]!, allCommands[1]!, ...allCommands.slice(3),
      ] }] },
      { paths: [{ ...geometry.paths[0]!, commands: [
        { kind: 'moveTo', point: { x: 1, y: 0 } },
      ] }] },
    ];
    for (const variant of variants) {
      expect(customGeometryEqual(
        normalized,
        normalizeCustomGeometry(variant, 'Custom geometry'),
      )).toBe(false);
    }
  });

  it('rejects malformed, ambiguous, formula, guide, handle, connection, and rect state', () => {
    const valid = canonical();
    const cases = [
      fixture(valid).replace(PRESENTATION_NAMESPACE, 'urn:wrong'),
      fixture(valid)
        .replace('<p:spPr>', '<q:spPr xmlns:q="urn:wrong">')
        .replace('</p:spPr>', '</q:spPr>'),
      fixture(valid).replace(valid, '<a:prstGeom prst="rect"/><a:custGeom/>'),
      fixture(valid).replace(valid, `${valid}${valid}`),
      fixture(valid).replace('<a:pathLst>', '<a:pathLst/><a:pathLst>'),
      fixture(valid).replace('<a:pathLst>', '<a:pathLst unsafe="1">'),
      fixture(valid).replace('<a:custGeom>', '<a:custGeom unsafe="1">'),
      fixture(valid).replace('<a:avLst/>', '<a:avLst><a:gd name="adj" fmla="val 1"/></a:avLst>'),
      fixture(valid).replace('<a:gdLst/>', '<a:gdLst><a:gd name="x" fmla="*/ w 1 2"/></a:gdLst>'),
      fixture(valid).replace('<a:ahLst/>', '<a:ahLst><a:ahXY/></a:ahLst>'),
      fixture(valid).replace('<a:cxnLst/>', '<a:cxnLst><a:cxn ang="0"/></a:cxnLst>'),
      fixture(valid).replace('l="l"', 'l="0"'),
      fixture(valid).replace('w="3657600"', 'w="0"'),
      fixture(valid).replace('w="3657600"', 'x:w="3657600" xmlns:x="urn:wrong"'),
      fixture(valid).replace('h="2743200"', 'h="unsafe"'),
      fixture(valid).replace('fill="norm"', 'fill="bad"'),
      fixture(valid).replace('stroke="1"', 'stroke="yes"'),
      fixture(valid).replace('x="914400"', 'x="wd2"'),
      fixture(valid)
        .replace('<a:moveTo>', '<x:moveTo xmlns:x="urn:wrong">')
        .replace('</a:moveTo>', '</x:moveTo>'),
      fixture(valid).replace('<a:pt x="0" y="0"/>', '<a:pt x="0" y="0"><a:pt x="1" y="1"/></a:pt>'),
      fixture(valid).replace('<a:close/>', '<a:close unsafe="1"/>'),
      fixture(valid).replace('<a:moveTo><a:pt x="0" y="0"/></a:moveTo>', '<a:lnTo><a:pt x="0" y="0"/></a:lnTo>'),
      fixture(valid).replace('</a:pathLst>', '<a:unknown/></a:pathLst>'),
      fixture(valid).replace('</a:path>', 'TEXT</a:path>'),
    ];
    for (const source of cases) {
      const { xml, shape } = parseShape(source);
      expect(readCustomGeometry(xml, shape), source).toBeUndefined();
      const before = xml.serialize();
      expect(() => replaceCustomGeometry(
        xml,
        shape,
        normalizeCustomGeometry(geometry, 'Custom geometry'),
        '/ppt/slides/slide1.xml',
      )).toThrow(ModelParseError);
      expect(xml.serialize()).toBe(before);
    }
  });

  it('performs exact no-op and prefix-preserving geometry-only replacement', () => {
    const normalized = normalizeCustomGeometry(geometry, 'Custom geometry');
    const same = parseShape(fixture(canonical()));
    const before = same.xml.serialize();
    expect(replaceCustomGeometry(
      same.xml,
      same.shape,
      normalized,
      '/ppt/slides/slide1.xml',
    )).toBe(false);
    expect(same.xml.serialize()).toBe(before);

    const replacement = normalizeCustomGeometry({
      paths: [{
        width: 20,
        height: 30,
        fill: 'none',
        stroke: false,
        commands: [
          { kind: 'moveTo', point: { x: 1, y: 2 } },
          { kind: 'lineTo', point: { x: 3, y: 4 } },
        ],
      }],
    }, 'Custom geometry');
    const alternate = parseShape(fixture(canonical('d:'), { drawingPrefix: 'd' }));
    expect(replaceCustomGeometry(
      alternate.xml,
      alternate.shape,
      replacement,
      '/ppt/slides/slide1.xml',
    )).toBe(true);
    const updated = alternate.xml.serialize();
    expect(updated).toContain(
      '<d:custGeom><d:avLst/><d:gdLst/><d:ahLst/><d:cxnLst/>' +
      '<d:rect l="l" t="t" r="r" b="b"/><d:pathLst>' +
      '<d:path w="20" h="30" fill="none" stroke="0">' +
      '<d:moveTo><d:pt x="1" y="2"/></d:moveTo>' +
      '<d:lnTo><d:pt x="3" y="4"/></d:lnTo></d:path></d:pathLst></d:custGeom>',
    );
    expect(updated).toContain('<d:xfrm/><d:solidFill/>');
    expect(updated).toContain('<d:ln/><d:effectLst/><d:extLst/>');
    expect(updated).toContain('<p:cNvPr id="2" name="Keep"/>');
    expect(updated).toContain('<d:bodyPr/><d:p/>');
  });

  it('adds a local namespace declaration when the geometry owns its prefix binding', () => {
    const source = fixture(canonical()).replace(
      ` xmlns:a="${DRAWING_NAMESPACE}"`,
      '',
    ).replace('<a:custGeom>', `<a:custGeom xmlns:a="${DRAWING_NAMESPACE}">`);
    const { xml, shape } = parseShape(source);
    const replacement = normalizeCustomGeometry({
      paths: [{ width: 1, height: 1, commands: [] }],
    }, 'Custom geometry');
    expect(replaceCustomGeometry(
      xml,
      shape,
      replacement,
      '/ppt/slides/slide1.xml',
    )).toBe(true);
    expect(xml.serialize()).toContain(
      `<a:custGeom xmlns:a="${DRAWING_NAMESPACE}"><a:avLst/>`,
    );
  });
});
