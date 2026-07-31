import { describe, expect, it } from 'vitest';
import { LosslessXmlDocument } from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
import type {
  AddCustomShapeOptions,
  CustomGeometry,
  CustomGeometryConnectionSite,
  CustomGeometryCommand,
  CustomGeometryFormula,
  CustomGeometryGuide,
  CustomGeometryHandle,
  CustomGeometryPolarHandle,
  CustomGeometryValue,
  CustomGeometryXyHandle,
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

const allFormulaGuides: readonly CustomGeometryGuide[] = [
  { name: 'gVal', formula: { operator: 'val', operands: [25_000] } },
  { name: 'gAbs', formula: { operator: 'abs', operands: [-25_000] } },
  { name: 'gSqrt', formula: { operator: 'sqrt', operands: [100] } },
  { name: 'gAt2', formula: { operator: 'at2', operands: ['h', 'w'] } },
  { name: 'gCos', formula: { operator: 'cos', operands: ['w', 'gAt2'] } },
  { name: 'gMax', formula: { operator: 'max', operands: ['w', 'h'] } },
  { name: 'gMin', formula: { operator: 'min', operands: ['w', 'h'] } },
  { name: 'gSin', formula: { operator: 'sin', operands: ['h', 'cd4'] } },
  { name: 'gTan', formula: { operator: 'tan', operands: ['ss', 'cd8'] } },
  { name: 'gMul', formula: { operator: '*/', operands: ['w', 'adj1', 100_000] } },
  { name: 'gAdd', formula: { operator: '+-', operands: ['h', 0, 'gMul'] } },
  { name: 'gDiv', formula: { operator: '+/', operands: ['w', 'h', 2] } },
  { name: 'gIf', formula: { operator: '?:', operands: ['adj1', 'gMul', 'gAdd'] } },
  { name: 'gCat', formula: { operator: 'cat2', operands: ['w', 'h', 'gAt2'] } },
  { name: 'gMod', formula: { operator: 'mod', operands: ['w', 'h', 0] } },
  { name: 'gPin', formula: { operator: 'pin', operands: [0, 'adj1', 100_000] } },
  { name: 'gSat', formula: { operator: 'sat2', operands: ['h', 'w', 'gAt2'] } },
];

const formulaGeometry: CustomGeometry = {
  adjustments: [{ name: 'adj1', formula: { operator: 'val', operands: [25_000] } }],
  guides: allFormulaGuides,
  paths: [{
    width: 100_000,
    height: 100_000,
    commands: [
      { kind: 'moveTo', point: { x: 'gMul', y: 0 } },
      { kind: 'lineTo', point: { x: 'r', y: 'gAdd' } },
      {
        kind: 'quadraticBezierTo',
        control: { x: 'gDiv', y: 'gMin' },
        end: { x: 'gMax', y: 'b' },
      },
      {
        kind: 'cubicBezierTo',
        control1: { x: 'gCos', y: 'gSin' },
        control2: { x: 'gCat', y: 'gSat' },
        end: { x: 'gPin', y: 'gMod' },
      },
      {
        kind: 'arcTo',
        widthRadius: 'gMul',
        heightRadius: 'hd2',
        startAngle: 'gAt2',
        sweepAngle: 'cd2',
      },
    ],
  }],
};

const handleGeometry: CustomGeometry = {
  adjustments: [
    { name: 'adjX', formula: { operator: 'val', operands: [25_000] } },
    { name: 'adjY', formula: { operator: 'val', operands: [50_000] } },
    { name: 'adjR', formula: { operator: 'val', operands: [30_000] } },
    { name: 'adjAng', formula: { operator: 'val', operands: [5_400_000] } },
  ],
  guides: [
    { name: 'x1', formula: { operator: '*/', operands: ['w', 'adjR', 100_000] } },
    { name: 'y1', formula: { operator: '*/', operands: ['h', 'adjR', 100_000] } },
  ],
  handles: [
    {
      kind: 'xy',
      position: { x: 'adjX', y: 'adjY' },
      xGuide: 'adjX',
      minX: 0,
      maxX: 100_000,
      yGuide: 'adjY',
      minY: 't',
      maxY: 'b',
    },
    {
      kind: 'polar',
      position: { x: 'x1', y: 'y1' },
      radiusGuide: 'adjR',
      minRadius: 0,
      maxRadius: 'ss',
      angleGuide: 'adjAng',
      minAngle: 0,
      maxAngle: 'cd',
    },
  ],
  paths: [{
    width: 100_000,
    height: 100_000,
    commands: [
      { kind: 'moveTo', point: { x: 'adjX', y: 0 } },
      { kind: 'lineTo', point: { x: 'r', y: 'adjY' } },
      { kind: 'close' },
    ],
  }],
};

const connectionGeometry: CustomGeometry = {
  connectionSites: [
    { angle: 0, position: { x: 'hc', y: 't' } },
    { angle: 'adjAng', position: { x: 'r', y: 60_000 } },
    { angle: -5_400_000, position: { x: 25_000, y: 100_000 } },
  ],
  paths: [{
    width: 100_000,
    height: 100_000,
    commands: [
      { kind: 'moveTo', point: { x: 0, y: 0 } },
      { kind: 'lineTo', point: { x: 100_000, y: 100_000 } },
    ],
  }],
};

const publicOptions: AddCustomShapeOptions = { name: 'Custom', x: inches(1) };
void publicOptions;
// @ts-expect-error preset-only adjustments are not custom-shape options
const invalidOptions: AddCustomShapeOptions = { adjustments: [] };
void invalidOptions;
const tokenCommand: CustomGeometryCommand = { kind: 'moveTo', point: { x: 'x1', y: 2 } };
void tokenCommand;
// @ts-expect-error custom geometry coordinates are number or string values
const invalidCommand: CustomGeometryCommand = { kind: 'moveTo', point: { x: false, y: 2 } };
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
const unaryFormula: CustomGeometryFormula = { operator: 'sqrt', operands: ['g1'] };
const binaryFormula: CustomGeometryFormula = { operator: 'max', operands: [1, 'g1'] };
const ternaryFormula: CustomGeometryFormula = { operator: 'pin', operands: [0, 'g1', 100_000] };
void [unaryFormula, binaryFormula, ternaryFormula];
// @ts-expect-error formula arity is encoded in the readonly tuple
const invalidFormulaArity: CustomGeometryFormula = { operator: 'val', operands: [1, 2] };
void invalidFormulaArity;
// @ts-expect-error formula operators are closed
const invalidFormulaOperator: CustomGeometryFormula = { operator: 'sum', operands: [1, 2, 3] };
void invalidFormulaOperator;
const xyHandle: CustomGeometryXyHandle = {
  kind: 'xy',
  position: { x: 0, y: 'vc' },
  xGuide: 'adjX',
};
const polarHandle: CustomGeometryPolarHandle = {
  kind: 'polar',
  position: { x: 'x1', y: 'y1' },
  angleGuide: 'adjAng',
};
const publicHandles: readonly CustomGeometryHandle[] = [xyHandle, polarHandle];
void publicHandles;
// @ts-expect-error adjustment handles require a position
const missingHandlePosition: CustomGeometryHandle = { kind: 'polar' };
void missingHandlePosition;
const invalidXyHandle: CustomGeometryXyHandle = {
  kind: 'xy',
  position: { x: 0, y: 0 },
  // @ts-expect-error XY handles do not accept polar guide fields
  radiusGuide: 'adjR',
};
void invalidXyHandle;
const publicConnectionSite: CustomGeometryConnectionSite = {
  angle: 'cd4',
  position: { x: 'r', y: 'vc' },
};
void publicConnectionSite;
// @ts-expect-error custom geometry connection sites require an angle
const missingConnectionAngle: CustomGeometryConnectionSite = { position: { x: 0, y: 0 } };
void missingConnectionAngle;

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

  it('copies and recursively freezes all guide formulas and token path values', () => {
    const mutable = structuredClone(formulaGeometry) as unknown as {
      adjustments: Array<{
        name: string;
        formula: { operator: string; operands: CustomGeometryValue[] };
      }>;
      guides: Array<{
        name: string;
        formula: { operator: string; operands: CustomGeometryValue[] };
      }>;
      paths: Array<{
        width: number;
        height: number;
        commands: Array<CustomGeometryCommand>;
      }>;
    };
    const normalized = normalizeCustomGeometry(mutable, 'Custom geometry');
    expect(normalized).toEqual(formulaGeometry);
    expect(Object.isFrozen(normalized)).toBe(true);
    expect(Object.isFrozen(normalized.adjustments)).toBe(true);
    expect(Object.isFrozen(normalized.adjustments?.[0])).toBe(true);
    expect(Object.isFrozen(normalized.adjustments?.[0]?.formula)).toBe(true);
    expect(Object.isFrozen(normalized.adjustments?.[0]?.formula.operands)).toBe(true);
    expect(Object.isFrozen(normalized.guides)).toBe(true);
    expect(normalized.guides?.every((guide) =>
      Object.isFrozen(guide)
      && Object.isFrozen(guide.formula)
      && Object.isFrozen(guide.formula.operands))).toBe(true);

    mutable.adjustments[0]!.name = 'changed';
    mutable.adjustments[0]!.formula.operands[0] = 99;
    mutable.guides.splice(0);
    mutable.paths[0]!.commands.splice(0);
    expect(normalized).toEqual(formulaGeometry);

    const empty = normalizeCustomGeometry({
      adjustments: [],
      guides: [],
      paths: [{ width: 1, height: 1, commands: [] }],
    }, 'Custom geometry');
    expect(Object.hasOwn(empty, 'adjustments')).toBe(false);
    expect(Object.hasOwn(empty, 'guides')).toBe(false);
  });

  it('accepts every guide formula operator at its exact arity', () => {
    const normalized = normalizeCustomGeometry(formulaGeometry, 'Custom geometry');
    expect(normalized.guides?.map(({ formula }) => formula.operator)).toEqual([
      'val', 'abs', 'sqrt', 'at2', 'cos', 'max', 'min', 'sin', 'tan',
      '*/', '+-', '+/', '?:', 'cat2', 'mod', 'pin', 'sat2',
    ]);
  });

  it('copies and recursively freezes ordered XY and polar handles', () => {
    const mutable = structuredClone(handleGeometry) as unknown as {
      handles: Array<{
        kind: string;
        position: { x: CustomGeometryValue; y: CustomGeometryValue };
        minX?: CustomGeometryValue;
        maxAngle?: CustomGeometryValue;
      }>;
      paths: Array<{ commands: CustomGeometryCommand[] }>;
    };
    const normalized = normalizeCustomGeometry(mutable, 'Custom geometry');
    expect(normalized).toEqual(handleGeometry);
    expect(Object.isFrozen(normalized)).toBe(true);
    expect(Object.isFrozen(normalized.handles)).toBe(true);
    expect(normalized.handles?.every((handle) =>
      Object.isFrozen(handle) && Object.isFrozen(handle.position))).toBe(true);

    mutable.handles[0]!.position.x = 99;
    mutable.handles[0]!.minX = -1;
    mutable.handles[1]!.maxAngle = 1;
    mutable.handles.reverse();
    mutable.paths[0]!.commands.splice(0);
    expect(normalized).toEqual(handleGeometry);

    const empty = normalizeCustomGeometry({
      handles: [],
      paths: [{ width: 1, height: 1, commands: [] }],
    }, 'Custom geometry');
    expect(Object.hasOwn(empty, 'handles')).toBe(false);
  });

  it('preserves independent optional handle fields without semantic inference', () => {
    const candidate: CustomGeometry = {
      handles: [
        { kind: 'xy', position: { x: -1, y: 0 }, minX: 10, maxX: -10 },
        { kind: 'xy', position: { x: 'l', y: 'b' }, yGuide: 'adjY', maxY: 'h' },
        { kind: 'polar', position: { x: 'hc', y: 'vc' }, minRadius: -1 },
        { kind: 'polar', position: { x: 0, y: 0 }, angleGuide: 'adjAng', maxAngle: 'cd' },
      ],
      paths: [{ width: 1, height: 1, commands: [] }],
    };
    expect(normalizeCustomGeometry(candidate, 'Custom geometry')).toEqual(candidate);
  });

  it('copies and recursively freezes ordered connection sites', () => {
    const mutable = structuredClone(connectionGeometry) as unknown as {
      connectionSites: Array<{
        angle: CustomGeometryValue;
        position: { x: CustomGeometryValue; y: CustomGeometryValue };
      }>;
    };
    const normalized = normalizeCustomGeometry(mutable, 'Custom geometry');
    expect(normalized).toEqual(connectionGeometry);
    expect(Object.isFrozen(normalized.connectionSites)).toBe(true);
    expect(normalized.connectionSites?.every((site) =>
      Object.isFrozen(site) && Object.isFrozen(site.position))).toBe(true);
    mutable.connectionSites[0]!.angle = 1;
    mutable.connectionSites[0]!.position.x = 1;
    mutable.connectionSites.reverse();
    expect(normalized).toEqual(connectionGeometry);

    const empty = normalizeCustomGeometry({
      connectionSites: [],
      paths: [{ width: 1, height: 1, commands: [] }],
    }, 'Custom geometry');
    expect(Object.hasOwn(empty, 'connectionSites')).toBe(false);
  });

  it('rejects unsafe connection site state without semantic inference', () => {
    const wrap = (connectionSites: unknown): unknown => ({
      connectionSites,
      paths: [{ width: 1, height: 1, commands: [] }],
    });
    for (const site of [
      null,
      {},
      { position: { x: 0, y: 0 } },
      { angle: 0 },
      { angle: 0.5, position: { x: 0, y: 0 } },
      { angle: Number.MAX_SAFE_INTEGER + 1, position: { x: 0, y: 0 } },
      { angle: '', position: { x: 0, y: 0 } },
      { angle: '1', position: { x: 0, y: 0 } },
      { angle: 'two words', position: { x: 0, y: 0 } },
      { angle: '\u0000', position: { x: 0, y: 0 } },
      { angle: 0, position: { x: 0, y: 0 }, extra: true },
      { angle: 0, position: { x: 0, y: 0 }, [Symbol('unsafe')]: true },
    ]) expect(() => normalizeCustomGeometry(wrap([site]), 'Custom geometry')).toThrow();
    const sparse: unknown[] = [];
    sparse.length = 1;
    expect(() => normalizeCustomGeometry(wrap(sparse), 'Custom geometry')).toThrow(/dense/);
    class ConnectionSites extends Array<unknown> {}
    expect(() => normalizeCustomGeometry(wrap(new ConnectionSites()), 'Custom geometry'))
      .toThrow(/ordinary array/);
    let calls = 0;
    const accessor = Object.defineProperties({}, {
      angle: {
        enumerable: true,
        get() {
          calls += 1;
          return 0;
        },
      },
      position: { enumerable: true, value: { x: 0, y: 0 } },
    });
    expect(() => normalizeCustomGeometry(wrap([accessor]), 'Custom geometry'))
      .toThrow(/data property/);
    expect(calls).toBe(0);
    expect(normalizeCustomGeometry({
      connectionSites: [
        { angle: -1, position: { x: -1, y: 0 } },
        { angle: 'cd', position: { x: 'hc', y: 'vc' } },
        { angle: -1, position: { x: -1, y: 0 } },
      ],
      paths: [{ width: 1, height: 1, commands: [] }],
    }, 'Custom geometry').connectionSites).toHaveLength(3);
  });

  it('rejects unsafe handle containers, kinds, fields, values, and accessors', () => {
    const wrap = (handles: unknown): unknown => ({
      handles,
      paths: [{ width: 1, height: 1, commands: [] }],
    });
    for (const handle of [
      null,
      false,
      {},
      { kind: 'xy' },
      { kind: 'unknown', position: { x: 0, y: 0 } },
      { kind: 'xy', position: { x: 0, y: 0 }, radiusGuide: 'adjR' },
      { kind: 'polar', position: { x: 0, y: 0 }, xGuide: 'adjX' },
      { kind: 'xy', position: { x: 0, y: 0 }, xGuide: '' },
      { kind: 'xy', position: { x: 0, y: 0 }, xGuide: 'two words' },
      { kind: 'xy', position: { x: 0, y: 0 }, xGuide: '1' },
      { kind: 'polar', position: { x: 0, y: 0 }, angleGuide: '\u0000' },
      { kind: 'xy', position: { x: 0, y: 0 }, maxX: Number.MAX_SAFE_INTEGER + 1 },
      { kind: 'polar', position: { x: 0.5, y: 0 } },
    ]) expect(() => normalizeCustomGeometry(wrap([handle]), 'Custom geometry')).toThrow();

    const sparse: unknown[] = [];
    sparse.length = 1;
    expect(() => normalizeCustomGeometry(wrap(sparse), 'Custom geometry')).toThrow(/dense/);
    class Handles extends Array<unknown> {}
    expect(() => normalizeCustomGeometry(wrap(new Handles()), 'Custom geometry'))
      .toThrow(/ordinary array/);

    let calls = 0;
    const accessor = Object.defineProperties({}, {
      kind: {
        enumerable: true,
        get() {
          calls += 1;
          return 'xy';
        },
      },
      position: { enumerable: true, value: { x: 0, y: 0 } },
    });
    expect(() => normalizeCustomGeometry(wrap([accessor]), 'Custom geometry'))
      .toThrow(/data property/);
    expect(calls).toBe(0);

    expect(() => normalizeCustomGeometry(wrap([{
      kind: 'xy',
      position: { x: 0, y: 0 },
      [Symbol('unsafe')]: true,
    }]), 'Custom geometry')).toThrow(/unsupported property/);
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

  it('rejects malformed guide formulas, duplicate names, and invalid value tokens', () => {
    const wrapGuides = (guides: unknown, adjustments: unknown = undefined) => ({
      ...(adjustments === undefined ? {} : { adjustments }),
      guides,
      paths: [{ width: 1, height: 1, commands: [] }],
    });
    for (const candidate of [
      wrapGuides(null),
      wrapGuides([null]),
      wrapGuides([{ name: '', formula: { operator: 'val', operands: [1] } }]),
      wrapGuides([{ name: 'two words', formula: { operator: 'val', operands: [1] } }]),
      wrapGuides([{ name: '1', formula: { operator: 'val', operands: [1] } }]),
      wrapGuides([{ name: 'bad\u0000', formula: { operator: 'val', operands: [1] } }]),
      wrapGuides([{ name: 'g1', formula: { operator: 'unknown', operands: [1] } }]),
      wrapGuides([{ name: 'g1', formula: { operator: 'val', operands: [] } }]),
      wrapGuides([{ name: 'g1', formula: { operator: 'val', operands: [1, 2] } }]),
      wrapGuides([{ name: 'g1', formula: { operator: 'max', operands: [1] } }]),
      wrapGuides([{ name: 'g1', formula: { operator: 'pin', operands: [1, 2] } }]),
      wrapGuides([{ name: 'g1', formula: { operator: 'val', operands: [1], extra: true } }]),
      wrapGuides([{ name: 'g1', formula: { operator: 'val', operands: [''] } }]),
      wrapGuides([{ name: 'g1', formula: { operator: 'val', operands: ['two words'] } }]),
      wrapGuides([{ name: 'g1', formula: { operator: 'val', operands: ['+1'] } }]),
      wrapGuides([{ name: 'g1', formula: { operator: 'val', operands: ['bad\u0000'] } }]),
      wrapGuides([{ name: 'g1', formula: { operator: 'val', operands: [1.5] } }]),
      wrapGuides([{ name: 'g1', formula: { operator: 'val', operands: [Number.MAX_SAFE_INTEGER + 1] } }]),
      wrapGuides([
        { name: 'g1', formula: { operator: 'val', operands: [1] } },
        { name: 'g1', formula: { operator: 'val', operands: [2] } },
      ]),
      wrapGuides(
        [{ name: 'g1', formula: { operator: 'val', operands: [1] } }],
        [{ name: 'g1', formula: { operator: 'val', operands: [2] } }],
      ),
      {
        paths: [{
          width: 1,
          height: 1,
          commands: [{ kind: 'moveTo', point: { x: 'two words', y: 0 } }],
        }],
      },
      {
        paths: [{
          width: 1,
          height: 1,
          commands: [{
            kind: 'moveTo',
            point: { x: 0, y: 0 },
          }, {
            kind: 'arcTo',
            widthRadius: '1',
            heightRadius: 'hd2',
            startAngle: 0,
            sweepAngle: 'cd2',
          }],
        }],
      },
    ]) expect(() => normalizeCustomGeometry(candidate, 'Custom geometry')).toThrow();

    const sparseOperands = Array(1);
    expect(() => normalizeCustomGeometry(wrapGuides([{
      name: 'g1',
      formula: { operator: 'val', operands: sparseOperands },
    }]), 'Custom geometry')).toThrow(/dense array/);

    class Operands<T> extends Array<T> {}
    expect(() => normalizeCustomGeometry(wrapGuides([{
      name: 'g1',
      formula: { operator: 'val', operands: new Operands(1) },
    }]), 'Custom geometry')).toThrow(/ordinary array/);

    let calls = 0;
    const formulaAccessor = Object.defineProperty({}, 'operator', {
      enumerable: true,
      get() {
        calls += 1;
        return 'val';
      },
    });
    Object.defineProperty(formulaAccessor, 'operands', { enumerable: true, value: [1] });
    expect(() => normalizeCustomGeometry(wrapGuides([{
      name: 'g1',
      formula: formulaAccessor,
    }]), 'Custom geometry')).toThrow(/data property/);
    expect(calls).toBe(0);
  });
});

describe('custom geometry OOXML codec', () => {
  it('renders every command, multiple paths, and direct path flags canonically', () => {
    const normalized = normalizeCustomGeometry(geometry, 'Custom geometry');
    expect(renderCustomGeometry(normalized, 'a:')).toBe(canonical());
    expect(renderCustomGeometry(normalized, '')).toBe(canonical(''));
  });

  it('renders guide formulas and token path values canonically with XML escaping', () => {
    const normalized = normalizeCustomGeometry(formulaGeometry, 'Custom geometry');
    const rendered = renderCustomGeometry(normalized, 'a:');
    expect(rendered).toContain(
      '<a:avLst><a:gd name="adj1" fmla="val 25000"/></a:avLst>',
    );
    expect(rendered).toContain(
      '<a:gd name="gMul" fmla="*/ w adj1 100000"/>',
    );
    expect(rendered).toContain(
      '<a:moveTo><a:pt x="gMul" y="0"/></a:moveTo>',
    );
    expect(rendered).toContain(
      '<a:arcTo wR="gMul" hR="hd2" stAng="gAt2" swAng="cd2"/>',
    );

    const escaped = normalizeCustomGeometry({
      guides: [{ name: 'x&1', formula: { operator: 'val', operands: ['x&1'] } }],
      paths: [{
        width: 1,
        height: 1,
        commands: [{ kind: 'moveTo', point: { x: 'x&1', y: 0 } }],
      }],
    }, 'Custom geometry');
    expect(renderCustomGeometry(escaped, 'd:')).toContain(
      '<d:gd name="x&amp;1" fmla="val x&amp;1"/>',
    );
    expect(renderCustomGeometry(escaped, 'd:')).toContain(
      '<d:pt x="x&amp;1" y="0"/>',
    );
  });

  it('renders ordered XY and polar handles with exact optional attributes', () => {
    const rendered = renderCustomGeometry(
      normalizeCustomGeometry(handleGeometry, 'Custom geometry'),
      'a:',
    );
    expect(rendered).toContain(
      '<a:ahLst>' +
      '<a:ahXY gdRefX="adjX" minX="0" maxX="100000" gdRefY="adjY" minY="t" maxY="b">' +
      '<a:pos x="adjX" y="adjY"/></a:ahXY>' +
      '<a:ahPolar gdRefR="adjR" minR="0" maxR="ss" gdRefAng="adjAng" minAng="0" maxAng="cd">' +
      '<a:pos x="x1" y="y1"/></a:ahPolar>' +
      '</a:ahLst>',
    );

    const escaped: CustomGeometry = {
      handles: [{
        kind: 'xy',
        position: { x: 'x&1', y: 'y&1' },
        xGuide: 'x&1',
        maxX: 'r&1',
      }],
      paths: [{ width: 1, height: 1, commands: [] }],
    };
    expect(renderCustomGeometry(
      normalizeCustomGeometry(escaped, 'Custom geometry'),
      'd:',
    )).toContain(
      '<d:ahXY gdRefX="x&amp;1" maxX="r&amp;1">' +
      '<d:pos x="x&amp;1" y="y&amp;1"/></d:ahXY>',
    );
  });

  it('renders ordered custom geometry connection sites', () => {
    const rendered = renderCustomGeometry(
      normalizeCustomGeometry(connectionGeometry, 'Custom geometry'),
      'a:',
    );
    expect(rendered).toContain(
      '<a:cxnLst>' +
      '<a:cxn ang="0"><a:pos x="hc" y="t"/></a:cxn>' +
      '<a:cxn ang="adjAng"><a:pos x="r" y="60000"/></a:cxn>' +
      '<a:cxn ang="-5400000"><a:pos x="25000" y="100000"/></a:cxn>' +
      '</a:cxnLst>',
    );
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

  it('reads guide formulas and token path values with semantic lexical normalization', () => {
    const normalized = normalizeCustomGeometry(formulaGeometry, 'Custom geometry');
    const formulaXml = renderCustomGeometry(normalized, 'a:');
    const alternateXml = renderCustomGeometry(normalized, 'd:');
    const lexicalXml = formulaXml
      .replace('fmla="val 25000"', 'fmla="  val   +25000  "')
      .replace('fmla="*/ w adj1 100000"', 'fmla="  */   w  adj1   +100000 "')
      .replace('x="gMul" y="0"', 'x="gMul" y="+0"');
    for (const source of [
      fixture(formulaXml),
      fixture(alternateXml, { drawingPrefix: 'd' }),
      fixture(lexicalXml),
    ]) {
      const { xml, shape } = parseShape(source);
      expect(readCustomGeometry(xml, shape)).toEqual(formulaGeometry);
    }

    const lexical = parseShape(fixture(lexicalXml));
    const before = lexical.xml.serialize();
    expect(replaceCustomGeometry(
      lexical.xml,
      lexical.shape,
      normalized,
      '/ppt/slides/slide1.xml',
    )).toBe(false);
    expect(lexical.xml.serialize()).toBe(before);

    const escapedGeometry: CustomGeometry = {
      guides: [{ name: 'x&1', formula: { operator: 'val', operands: ['x&1'] } }],
      paths: [{
        width: 1,
        height: 1,
        commands: [{ kind: 'moveTo', point: { x: 'x&1', y: 0 } }],
      }],
    };
    const escaped = parseShape(fixture(renderCustomGeometry(
      normalizeCustomGeometry(escapedGeometry, 'Custom geometry'),
      'a:',
    )));
    expect(readCustomGeometry(escaped.xml, escaped.shape)).toEqual(escapedGeometry);

    const positionOnlyGeometry: CustomGeometry = {
      handles: [
        { kind: 'xy', position: { x: 0, y: 'vc' } },
        { kind: 'polar', position: { x: 'hc', y: 0 } },
      ],
      paths: [{ width: 1, height: 1, commands: [] }],
    };
    const positionOnly = parseShape(fixture(renderCustomGeometry(
      normalizeCustomGeometry(positionOnlyGeometry, 'Custom geometry'),
      'a:',
    )));
    expect(readCustomGeometry(positionOnly.xml, positionOnly.shape))
      .toEqual(positionOnlyGeometry);
  });

  it('reads ordered XY and polar handles with semantic lexical normalization', () => {
    const normalized = normalizeCustomGeometry(handleGeometry, 'Custom geometry');
    const handleXml = renderCustomGeometry(normalized, 'a:');
    const alternateXml = renderCustomGeometry(normalized, 'd:');
    const lexicalXml = handleXml
      .replace('minX="0"', 'minX="+0"')
      .replace('maxX="100000"', 'maxX="+100000"')
      .replace('minAng="0"', 'minAng="+0"');
    for (const source of [
      fixture(handleXml),
      fixture(alternateXml, { drawingPrefix: 'd' }),
      fixture(lexicalXml),
    ]) {
      const { xml, shape } = parseShape(source);
      expect(readCustomGeometry(xml, shape)).toEqual(handleGeometry);
    }

    const lexical = parseShape(fixture(lexicalXml));
    const before = lexical.xml.serialize();
    expect(replaceCustomGeometry(
      lexical.xml,
      lexical.shape,
      normalized,
      '/ppt/slides/slide1.xml',
    )).toBe(false);
    expect(lexical.xml.serialize()).toBe(before);

    const escapedGeometry: CustomGeometry = {
      handles: [{
        kind: 'polar',
        position: { x: 'x&1', y: 'y&1' },
        radiusGuide: 'r&1',
        maxRadius: 'ss&1',
      }],
      paths: [{ width: 1, height: 1, commands: [] }],
    };
    const escaped = parseShape(fixture(renderCustomGeometry(
      normalizeCustomGeometry(escapedGeometry, 'Custom geometry'),
      'a:',
    )));
    expect(readCustomGeometry(escaped.xml, escaped.shape)).toEqual(escapedGeometry);
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

  it('distinguishes guide list presence, order, names, operators, and operands', () => {
    const normalized = normalizeCustomGeometry(formulaGeometry, 'Custom geometry');
    const { adjustments: _adjustments, ...withoutAdjustments } = formulaGeometry;
    const variants: CustomGeometry[] = [
      withoutAdjustments,
      { ...formulaGeometry, guides: [...allFormulaGuides].reverse() },
      {
        ...formulaGeometry,
        guides: [
          { ...allFormulaGuides[0]!, name: 'renamed' },
          ...allFormulaGuides.slice(1),
        ],
      },
      {
        ...formulaGeometry,
        guides: [
          { name: 'gVal', formula: { operator: 'abs', operands: [25_000] } },
          ...allFormulaGuides.slice(1),
        ],
      },
      {
        ...formulaGeometry,
        guides: [
          { name: 'gVal', formula: { operator: 'val', operands: [25_001] } },
          ...allFormulaGuides.slice(1),
        ],
      },
      {
        ...formulaGeometry,
        paths: [{
          ...formulaGeometry.paths[0]!,
          commands: [{ kind: 'moveTo', point: { x: 'gAdd', y: 0 } }],
        }],
      },
    ];
    for (const variant of variants) {
      expect(customGeometryEqual(
        normalized,
        normalizeCustomGeometry(variant, 'Custom geometry'),
      )).toBe(false);
    }
  });

  it('distinguishes handle presence, order, kind, position, and optional fields', () => {
    const normalized = normalizeCustomGeometry(handleGeometry, 'Custom geometry');
    const { handles: _handles, ...withoutHandles } = handleGeometry;
    const xy = handleGeometry.handles?.[0];
    const polar = handleGeometry.handles?.[1];
    if (!xy || xy.kind !== 'xy' || !polar || polar.kind !== 'polar') {
      throw new Error('Handle fixture is invalid');
    }
    const { minX: _minX, ...xyWithoutMin } = xy;
    const variants: CustomGeometry[] = [
      withoutHandles,
      { ...handleGeometry, handles: [...handleGeometry.handles!].reverse() },
      { ...handleGeometry, handles: [{ ...xy, position: { x: 'adjY', y: 'adjY' } }, polar] },
      { ...handleGeometry, handles: [xyWithoutMin, polar] },
      { ...handleGeometry, handles: [{ ...xy, maxX: 99_999 }, polar] },
      {
        ...handleGeometry,
        handles: [{ kind: 'polar', position: xy.position }, polar],
      },
    ];
    for (const variant of variants) {
      expect(customGeometryEqual(
        normalized,
        normalizeCustomGeometry(variant, 'Custom geometry'),
      )).toBe(false);
    }
  });

  it('distinguishes connection site presence, order, angle, and position', () => {
    const normalized = normalizeCustomGeometry(connectionGeometry, 'Custom geometry');
    const { connectionSites: _sites, ...withoutSites } = connectionGeometry;
    const sites = connectionGeometry.connectionSites!;
    const variants: CustomGeometry[] = [
      withoutSites,
      { ...connectionGeometry, connectionSites: [...sites].reverse() },
      { ...connectionGeometry, connectionSites: sites.slice(0, 2) },
      { ...connectionGeometry, connectionSites: [{ ...sites[0]!, angle: 1 }, ...sites.slice(1)] },
      {
        ...connectionGeometry,
        connectionSites: [
          { ...sites[0]!, position: { ...sites[0]!.position, x: 'l' } },
          ...sites.slice(1),
        ],
      },
    ];
    for (const variant of variants) {
      expect(customGeometryEqual(
        normalized,
        normalizeCustomGeometry(variant, 'Custom geometry'),
      )).toBe(false);
    }
  });

  it('rejects malformed, ambiguous, guide, handle, connection, and rect state', () => {
    const valid = canonical();
    const validHandles = renderCustomGeometry(
      normalizeCustomGeometry(handleGeometry, 'Custom geometry'),
      'a:',
    );
    const xyPosition = '<a:pos x="adjX" y="adjY"/>';
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
      fixture(valid).replace('<a:avLst/>', '<a:avLst unsafe="1"/>'),
      fixture(valid).replace('<a:gdLst/>', '<a:gdLst/><a:gdLst/>'),
      fixture(valid).replace('<a:gdLst/>', '<a:gdLst><a:unknown/></a:gdLst>'),
      fixture(valid).replace('<a:gdLst/>', '<a:gdLst><x:gd xmlns:x="urn:wrong" name="x" fmla="val 1"/></a:gdLst>'),
      fixture(valid).replace('<a:gdLst/>', '<a:gdLst><a:gd fmla="val 1"/></a:gdLst>'),
      fixture(valid).replace('<a:gdLst/>', '<a:gdLst><a:gd name="x"/></a:gdLst>'),
      fixture(valid).replace('<a:gdLst/>', '<a:gdLst><a:gd name="x" fmla="val 1" extra="1"/></a:gdLst>'),
      fixture(valid).replace('<a:gdLst/>', '<a:gdLst><a:gd x:name="x" xmlns:x="urn:wrong" fmla="val 1"/></a:gdLst>'),
      fixture(valid).replace('<a:gdLst/>', '<a:gdLst><a:gd name="x" x:fmla="val 1" xmlns:x="urn:wrong"/></a:gdLst>'),
      fixture(valid).replace('<a:gdLst/>', '<a:gdLst><a:gd name="x" fmla="unknown 1"/></a:gdLst>'),
      fixture(valid).replace('<a:gdLst/>', '<a:gdLst><a:gd name="x" fmla="val"/></a:gdLst>'),
      fixture(valid).replace('<a:gdLst/>', '<a:gdLst><a:gd name="x" fmla="val 1 2"/></a:gdLst>'),
      fixture(valid).replace('<a:gdLst/>', '<a:gdLst><a:gd name="x" fmla="max 1"/></a:gdLst>'),
      fixture(valid).replace('<a:gdLst/>', '<a:gdLst><a:gd name="x" fmla="pin 1 2"/></a:gdLst>'),
      fixture(valid).replace('<a:gdLst/>', '<a:gdLst><a:gd name="x" fmla="val 9007199254740992"/></a:gdLst>'),
      fixture(valid).replace('<a:gdLst/>', '<a:gdLst><a:gd name="x" fmla="val two words"/></a:gdLst>'),
      fixture(valid).replace(
        '<a:avLst/><a:gdLst/>',
        '<a:avLst><a:gd name="same" fmla="val 1"/></a:avLst>' +
        '<a:gdLst><a:gd name="same" fmla="val 2"/></a:gdLst>',
      ),
      fixture(valid).replace('<a:gdLst/>', '<a:gdLst><a:gd name="x" fmla="val 1"><a:gd/></a:gd></a:gdLst>'),
      fixture(valid).replace('<a:ahLst/>', '<a:ahLst><a:ahXY/></a:ahLst>'),
      fixture(valid).replace('<a:cxnLst/>', '<a:cxnLst><a:cxn ang="0"/></a:cxnLst>'),
      fixture(valid).replace('l="l"', 'l="0"'),
      fixture(valid).replace('w="3657600"', 'w="0"'),
      fixture(valid).replace('w="3657600"', 'x:w="3657600" xmlns:x="urn:wrong"'),
      fixture(valid).replace('h="2743200"', 'h="unsafe"'),
      fixture(valid).replace('fill="norm"', 'fill="bad"'),
      fixture(valid).replace('stroke="1"', 'stroke="yes"'),
      fixture(valid).replace('x="914400"', 'x="two words"'),
      fixture(valid)
        .replace('<a:moveTo>', '<x:moveTo xmlns:x="urn:wrong">')
        .replace('</a:moveTo>', '</x:moveTo>'),
      fixture(valid).replace('<a:pt x="0" y="0"/>', '<a:pt x="0" y="0"><a:pt x="1" y="1"/></a:pt>'),
      fixture(valid).replace('<a:close/>', '<a:close unsafe="1"/>'),
      fixture(valid).replace('<a:moveTo><a:pt x="0" y="0"/></a:moveTo>', '<a:lnTo><a:pt x="0" y="0"/></a:lnTo>'),
      fixture(valid).replace('</a:pathLst>', '<a:unknown/></a:pathLst>'),
      fixture(valid).replace('</a:path>', 'TEXT</a:path>'),
      fixture(validHandles).replace('<a:ahLst>', '<a:ahLst/><a:ahLst>'),
      fixture(validHandles).replace('<a:ahLst>', '<a:ahLst unsafe="1">'),
      fixture(validHandles).replace('<a:ahLst>', '<a:ahLst>TEXT'),
      fixture(validHandles)
        .replace('<a:ahXY ', '<x:ahXY xmlns:x="urn:wrong" ')
        .replace('</a:ahXY>', '</x:ahXY>'),
      fixture(validHandles).replaceAll('a:ahXY', 'a:unknown'),
      fixture(validHandles).replace(xyPosition, ''),
      fixture(validHandles).replace(xyPosition, `${xyPosition}${xyPosition}`),
      fixture(validHandles).replace(
        xyPosition,
        '<x:pos xmlns:x="urn:wrong" x="adjX" y="adjY"/>',
      ),
      fixture(validHandles).replace('</a:ahXY>', '<a:pt x="0" y="0"/></a:ahXY>'),
      fixture(validHandles).replace('</a:ahXY>', 'TEXT</a:ahXY>'),
      fixture(validHandles).replace('<a:ahXY ', '<a:ahXY unsafe="1" '),
      fixture(validHandles).replace(
        '<a:ahXY gdRefX="adjX"',
        '<a:ahXY xmlns:x="urn:wrong" x:gdRefX="adjX"',
      ),
      fixture(validHandles).replace(xyPosition, '<a:pos y="adjY"/>'),
      fixture(validHandles).replace(
        xyPosition,
        '<a:pos xmlns:x="urn:wrong" x:x="adjX" y="adjY"/>',
      ),
      fixture(validHandles).replace(xyPosition, '<a:pos x="adjX" y="adjY" extra="1"/>'),
      fixture(validHandles).replace('minX="0"', 'minX="9007199254740992"'),
      fixture(validHandles).replace('minX="0"', 'minX="two words"'),
      fixture(validHandles).replace('gdRefX="adjX"', 'gdRefX=""'),
      fixture(validHandles).replace('minAng="0"', 'minAng="two words"'),
      fixture(validHandles).replace('gdRefR="adjR"', 'gdRefR=""'),
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

  it('converts supported preset geometry and rejects malformed preset ownership', () => {
    const normalized = normalizeCustomGeometry(geometry, 'Custom geometry');
    const source = fixture(
      '<d:prstGeom prst="blockArc"><d:avLst>' +
      '<d:gd name="adj1" fmla="val 16200000"/></d:avLst></d:prstGeom>',
      { drawingPrefix: 'd' },
    );
    const converted = parseShape(source);
    expect(replaceCustomGeometry(
      converted.xml,
      converted.shape,
      normalized,
      '/ppt/slides/slide1.xml',
    )).toBe(true);
    const updated = converted.xml.serialize();
    expect(updated).toContain(canonical('d:'));
    expect(updated).not.toContain('<d:prstGeom');
    expect(updated).not.toContain('name="adj1"');
    expect(updated).toContain('<d:xfrm/><d:solidFill/>');
    expect(updated).toContain('<d:ln/><d:effectLst/><d:extLst/>');
    expect(updated).toContain('<p:cNvPr id="2" name="Keep"/>');

    for (const malformed of [
      '<a:prstGeom/>',
      '<a:prstGeom prst="folderCorner"/>',
      '<a:prstGeom x:prst="rect" xmlns:x="urn:wrong"/>',
      '<a:prstGeom prst="rect"/><a:custGeom/>',
      '<a:prstGeom prst="rect"/><a:prstGeom prst="ellipse"/>',
    ]) {
      const candidate = parseShape(fixture(malformed));
      const before = candidate.xml.serialize();
      expect(() => replaceCustomGeometry(
        candidate.xml,
        candidate.shape,
        normalized,
        '/ppt/slides/slide1.xml',
      )).toThrow(ModelParseError);
      expect(candidate.xml.serialize()).toBe(before);
    }
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
