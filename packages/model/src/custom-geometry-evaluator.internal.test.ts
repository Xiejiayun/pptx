import { describe, expect, it } from 'vitest';
import {
  CustomGeometryEvaluationError,
  type CustomGeometry,
  type CustomGeometryEvaluationContext,
  type CustomGeometryFormula,
  type EvaluatedCustomGeometry,
  type EvaluatedCustomGeometryCommand,
  type EvaluatedCustomGeometryConnectionSite,
  type EvaluatedCustomGeometryGuide,
  type EvaluatedCustomGeometryHandle,
  type EvaluatedCustomGeometryPath,
  type EvaluatedCustomGeometryPoint,
  type EvaluatedCustomGeometryPolarHandle,
  type EvaluatedCustomGeometryTextRectangle,
  type EvaluatedCustomGeometryXyHandle,
} from './custom-geometry.js';
import {
  evaluateBuiltInGuide,
  evaluateFormulaValue,
  evaluateGuideEnvironment,
} from './custom-geometry-evaluator.internal.js';
import { evaluateCustomGeometry } from './custom-geometry-evaluator.js';

const OOXML_DEGREE = 60_000;

const evaluatedPoint: EvaluatedCustomGeometryPoint = { x: 1, y: 2 };
const evaluatedGuide: EvaluatedCustomGeometryGuide = { name: 'g1', value: 1 };
const evaluatedXyHandle: EvaluatedCustomGeometryXyHandle = {
  kind: 'xy',
  position: evaluatedPoint,
  xGuide: 'adj1',
  minX: 0,
  maxX: 100,
};
const evaluatedPolarHandle: EvaluatedCustomGeometryPolarHandle = {
  kind: 'polar',
  position: evaluatedPoint,
  radiusGuide: 'adj2',
  minRadius: 1,
  maxRadius: 100,
  angleGuide: 'adj3',
  minAngle: 0,
  maxAngle: 360 * OOXML_DEGREE,
};
const evaluatedHandle: EvaluatedCustomGeometryHandle = evaluatedXyHandle;
const evaluatedConnectionSite: EvaluatedCustomGeometryConnectionSite = {
  position: evaluatedPoint,
  angle: 0,
};
const evaluatedTextRectangle: EvaluatedCustomGeometryTextRectangle = {
  left: 0,
  top: 0,
  right: 100,
  bottom: 100,
};
const evaluatedCommand: EvaluatedCustomGeometryCommand = {
  kind: 'moveTo',
  point: evaluatedPoint,
};
const evaluatedPath: EvaluatedCustomGeometryPath = {
  width: 100,
  height: 100,
  commands: [evaluatedCommand],
};
const evaluatedGeometry: EvaluatedCustomGeometry = {
  context: { width: 100, height: 100 },
  adjustments: [evaluatedGuide],
  handles: [evaluatedHandle, evaluatedPolarHandle],
  connectionSites: [evaluatedConnectionSite],
  textRectangle: evaluatedTextRectangle,
  paths: [evaluatedPath],
};

void evaluatedGeometry;

describe('custom geometry scalar evaluation', () => {
  it('evaluates all 37 DrawingML built-in guides', () => {
    const context: CustomGeometryEvaluationContext = { width: 120, height: 80 };
    const expected = new Map<string, number>([
      ['3cd4', 270 * OOXML_DEGREE],
      ['3cd8', 135 * OOXML_DEGREE],
      ['5cd8', 225 * OOXML_DEGREE],
      ['7cd8', 315 * OOXML_DEGREE],
      ['b', 80],
      ['cd2', 180 * OOXML_DEGREE],
      ['cd4', 90 * OOXML_DEGREE],
      ['cd8', 45 * OOXML_DEGREE],
      ['hc', 60],
      ['h', 80],
      ['hd2', 40],
      ['hd3', 80 / 3],
      ['hd4', 20],
      ['hd5', 16],
      ['hd6', 80 / 6],
      ['hd8', 10],
      ['l', 0],
      ['ls', 120],
      ['r', 120],
      ['ss', 80],
      ['ssd2', 40],
      ['ssd4', 20],
      ['ssd6', 80 / 6],
      ['ssd8', 10],
      ['ssd16', 5],
      ['ssd32', 2.5],
      ['t', 0],
      ['vc', 40],
      ['w', 120],
      ['wd2', 60],
      ['wd3', 40],
      ['wd4', 30],
      ['wd5', 24],
      ['wd6', 20],
      ['wd8', 15],
      ['wd10', 12],
      ['wd32', 3.75],
    ]);

    expect(expected.size).toBe(37);
    for (const [token, value] of expected) {
      expect(evaluateBuiltInGuide(token, context), token).toBe(value);
    }
    expect(evaluateBuiltInGuide('unknown', context)).toBeUndefined();
  });

  it('preserves fractional built-in results for odd extents', () => {
    const context: CustomGeometryEvaluationContext = { width: 101, height: 67 };

    expect(evaluateBuiltInGuide('wd2', context)).toBe(50.5);
    expect(evaluateBuiltInGuide('hd3', context)).toBe(67 / 3);
    expect(evaluateBuiltInGuide('ssd32', context)).toBe(67 / 32);
  });

  it('evaluates all 17 DrawingML formula operators', () => {
    const values = new Map<string, number>([
      ['x', 3],
      ['y', 4],
      ['z', 5],
      ['rightAngle', 90 * OOXML_DEGREE],
    ]);
    const evaluate = (formula: CustomGeometryFormula): number => evaluateFormulaValue(
      formula,
      (value) => typeof value === 'number' ? value : values.get(value)!,
      'g1',
    );

    expect(evaluate({ operator: 'val', operands: ['x'] })).toBe(3);
    expect(evaluate({ operator: 'abs', operands: [-3] })).toBe(3);
    expect(evaluate({ operator: 'sqrt', operands: [144] })).toBe(12);
    expect(evaluate({ operator: 'at2', operands: [0, 1] })).toBe(90 * OOXML_DEGREE);
    expect(evaluate({ operator: 'cos', operands: [2, 0] })).toBe(2);
    expect(evaluate({ operator: 'max', operands: ['x', 'y'] })).toBe(4);
    expect(evaluate({ operator: 'min', operands: ['x', 'y'] })).toBe(3);
    expect(evaluate({ operator: 'sin', operands: [2, 'rightAngle'] })).toBeCloseTo(2);
    expect(evaluate({ operator: 'tan', operands: [2, 45 * OOXML_DEGREE] })).toBeCloseTo(2);
    expect(evaluate({ operator: '*/', operands: ['x', 'y', 2] })).toBe(6);
    expect(evaluate({ operator: '+-', operands: ['x', 'y', 2] })).toBe(5);
    expect(evaluate({ operator: '+/', operands: ['x', 'y', 2] })).toBe(3.5);
    expect(evaluate({ operator: '?:', operands: [0, 10, 20] })).toBe(20);
    expect(evaluate({ operator: 'cat2', operands: [10, 3, 4] })).toBeCloseTo(6);
    expect(evaluate({ operator: 'mod', operands: ['x', 'y', 12] })).toBe(13);
    expect(evaluate({ operator: 'pin', operands: [0, 12, 10] })).toBe(10);
    expect(evaluate({ operator: 'sat2', operands: [10, 3, 4] })).toBeCloseTo(8);
  });

  it('matches DrawingML zero and quadrant semantics', () => {
    const evaluate = (formula: CustomGeometryFormula): number => evaluateFormulaValue(
      formula,
      (value) => value as number,
      'g1',
    );

    expect(evaluate({ operator: '*/', operands: [3, 4, 0] })).toBe(0);
    expect(evaluate({ operator: '+/', operands: [3, 4, 0] })).toBe(0);
    expect(evaluate({ operator: '?:', operands: [-1, 10, 20] })).toBe(20);
    expect(evaluate({ operator: '?:', operands: [1, 10, 20] })).toBe(10);
    expect(evaluate({ operator: 'at2', operands: [-1, 0] })).toBe(180 * OOXML_DEGREE);
    expect(evaluate({ operator: 'at2', operands: [0, -1] })).toBe(-90 * OOXML_DEGREE);
    expect(evaluate({ operator: 'cat2', operands: [10, 0, 0] })).toBe(10);
    expect(evaluate({ operator: 'sat2', operands: [10, 0, 0] })).toBe(0);
    expect(Object.is(evaluate({ operator: 'val', operands: [-0] }), -0)).toBe(false);
  });

  it('resolves every formula operand in source order before applying the operator', () => {
    const resolved: (number | string)[] = [];
    const value = evaluateFormulaValue(
      { operator: '?:', operands: [1, 10, 20] },
      (operand) => {
        resolved.push(operand);
        return operand as number;
      },
      'conditional',
    );

    expect(value).toBe(10);
    expect(resolved).toEqual([1, 10, 20]);
  });

  it('rejects invalid domains and non-finite results', () => {
    const negativeSqrt = (): number => evaluateFormulaValue(
      { operator: 'sqrt', operands: [-1] },
      (value) => value as number,
      'negativeRoot',
    );
    const overflow = (): number => evaluateFormulaValue(
      { operator: '*/', operands: ['huge', 'huge', 1] },
      (value) => value === 'huge' ? Number.MAX_VALUE : value as number,
      'overflow',
    );

    expect(negativeSqrt).toThrow(CustomGeometryEvaluationError);
    expectEvaluationError(negativeSqrt, 'invalid-domain', 'negativeRoot');
    expect(overflow).toThrow(CustomGeometryEvaluationError);
    expectEvaluationError(overflow, 'non-finite-result', 'overflow');
  });
});

describe('custom geometry guide environment', () => {
  it('evaluates adjustments before guides in source order', () => {
    const geometry = geometryWithGuides(
      [{ name: 'adj', formula: { operator: 'val', operands: [25_000] } }],
      [
        { name: 'x1', formula: { operator: '*/', operands: ['w', 'adj', 100_000] } },
        { name: 'w', formula: { operator: 'val', operands: ['w'] } },
        { name: 'x2', formula: { operator: '+-', operands: ['w', 'x1', 0] } },
      ],
    );

    const environment = evaluateGuideEnvironment(
      geometry,
      { width: 200_000, height: 100_000 },
    );

    expect(environment.adjustments).toEqual([{ name: 'adj', value: 25_000 }]);
    expect(environment.guides).toEqual([
      { name: 'x1', value: 50_000 },
      { name: 'w', value: 200_000 },
      { name: 'x2', value: 250_000 },
    ]);
    expect(environment.resolve('x2', 'test value')).toBe(250_000);
    expect(environment.resolve('h', 'test value')).toBe(100_000);
    expect(environment.resolve(12.5, 'test value')).toBe(12.5);
    expect(Object.isFrozen(environment)).toBe(true);
    expect(Object.isFrozen(environment.adjustments)).toBe(true);
    expect(Object.isFrozen(environment.adjustments?.[0])).toBe(true);
    expect(Object.isFrozen(environment.guides)).toBe(true);
    expect(Object.isFrozen(environment.guides?.[0])).toBe(true);
  });

  it('lets an evaluated custom guide shadow a built-in for later formulas', () => {
    const environment = evaluateGuideEnvironment(
      geometryWithGuides(undefined, [
        { name: 'w', formula: { operator: 'val', operands: [123] } },
        { name: 'after', formula: { operator: 'val', operands: ['w'] } },
      ]),
      { width: 200_000, height: 100_000 },
    );

    expect(environment.guides).toEqual([
      { name: 'w', value: 123 },
      { name: 'after', value: 123 },
    ]);
    expect(environment.resolve('w', 'test value')).toBe(123);
  });

  it('uses a same-name built-in while defining its custom shadow', () => {
    const environment = evaluateGuideEnvironment(
      geometryWithGuides(undefined, [
        { name: 'w', formula: { operator: '+-', operands: ['w', 1, 0] } },
        { name: 'after', formula: { operator: 'val', operands: ['w'] } },
      ]),
      { width: 200_000, height: 100_000 },
    );

    expect(environment.guides).toEqual([
      { name: 'w', value: 200_001 },
      { name: 'after', value: 200_001 },
    ]);
  });

  it('distinguishes cycles, forward references, and unknown tokens', () => {
    const context: CustomGeometryEvaluationContext = { width: 100, height: 100 };
    const selfCycle = (): unknown => evaluateGuideEnvironment(
      geometryWithGuides(undefined, [
        { name: 'g1', formula: { operator: 'val', operands: ['g1'] } },
      ]),
      context,
    );
    const multiNodeCycle = (): unknown => evaluateGuideEnvironment(
      geometryWithGuides(undefined, [
        { name: 'g1', formula: { operator: 'val', operands: ['g2'] } },
        { name: 'g2', formula: { operator: 'val', operands: ['g3'] } },
        { name: 'g3', formula: { operator: 'val', operands: ['g1'] } },
      ]),
      context,
    );
    const forwardReference = (): unknown => evaluateGuideEnvironment(
      geometryWithGuides(undefined, [
        { name: 'g1', formula: { operator: 'val', operands: ['g2'] } },
        { name: 'g2', formula: { operator: 'val', operands: [1] } },
      ]),
      context,
    );
    const unknownToken = (): unknown => evaluateGuideEnvironment(
      geometryWithGuides(undefined, [
        { name: 'g1', formula: { operator: 'val', operands: ['missingGuide'] } },
      ]),
      context,
    );

    expectEvaluationError(selfCycle, 'cyclic-reference', 'g1', 'g1');
    expectEvaluationError(multiNodeCycle, 'cyclic-reference', 'g1', 'g2');
    expectEvaluationError(forwardReference, 'forward-reference', 'g1', 'g2');
    expectEvaluationError(unknownToken, 'unknown-token', 'g1', 'missingGuide');
  });

  it('rejects unknown tokens resolved outside guide formulas', () => {
    const environment = evaluateGuideEnvironment(
      geometryWithGuides(undefined, undefined),
      { width: 100, height: 100 },
    );
    const operation = (): number => environment.resolve('missing', 'Path point x');

    expectEvaluationError(operation, 'unknown-token', undefined, 'missing');
  });
});

describe('custom geometry tree evaluation', () => {
  it('resolves the complete custom geometry tree without scaling path coordinates', () => {
    const geometry: CustomGeometry = {
      adjustments: [{ name: 'adj', formula: { operator: 'val', operands: [25_000] } }],
      guides: [
        { name: 'x1', formula: { operator: '*/', operands: ['w', 'adj', 100_000] } },
        { name: 'y1', formula: { operator: '+-', operands: ['h', 0, 'adj'] } },
        { name: 'rad', formula: { operator: 'val', operands: [10_000] } },
      ],
      handles: [
        {
          kind: 'xy',
          position: { x: 'x1', y: 'vc' },
          xGuide: 'adj',
          minX: 0,
          maxX: 'r',
          minY: 't',
        },
        {
          kind: 'polar',
          position: { x: 'hc', y: 'vc' },
          radiusGuide: 'adj',
          minRadius: 1,
          maxRadius: 'ss',
          angleGuide: 'adj',
          minAngle: 0,
          maxAngle: 'cd2',
        },
      ],
      connectionSites: [
        { angle: 0, position: { x: 'l', y: 't' } },
        { angle: 0, position: { x: 'l', y: 't' } },
      ],
      textRectangle: { left: 'x1', top: 1_000, right: 'r', bottom: 'y1' },
      paths: [
        {
          width: 100_000,
          height: 100_000,
          fill: 'norm',
          stroke: false,
          extrusionOk: true,
          commands: [
            { kind: 'moveTo', point: { x: 'x1', y: 0 } },
            { kind: 'lineTo', point: { x: 'r', y: 'b' } },
            {
              kind: 'quadraticBezierTo',
              control: { x: 'hc', y: 't' },
              end: { x: 'r', y: 'vc' },
            },
            {
              kind: 'cubicBezierTo',
              control1: { x: 1, y: 2 },
              control2: { x: 'x1', y: 'y1' },
              end: { x: 'r', y: 'b' },
            },
            {
              kind: 'arcTo',
              widthRadius: 'rad',
              heightRadius: 'hd2',
              startAngle: 'cd4',
              sweepAngle: 'cd2',
            },
            { kind: 'close' },
          ],
        },
        { width: 1, height: 1, commands: [] },
      ],
    };

    const evaluated = evaluateCustomGeometry(
      geometry,
      { width: 200_000, height: 100_000 },
    );

    expect(evaluated).toEqual({
      context: { width: 200_000, height: 100_000 },
      adjustments: [{ name: 'adj', value: 25_000 }],
      guides: [
        { name: 'x1', value: 50_000 },
        { name: 'y1', value: 75_000 },
        { name: 'rad', value: 10_000 },
      ],
      handles: [
        {
          kind: 'xy',
          position: { x: 50_000, y: 50_000 },
          xGuide: 'adj',
          minX: 0,
          maxX: 200_000,
          minY: 0,
        },
        {
          kind: 'polar',
          position: { x: 100_000, y: 50_000 },
          radiusGuide: 'adj',
          minRadius: 1,
          maxRadius: 100_000,
          angleGuide: 'adj',
          minAngle: 0,
          maxAngle: 180 * OOXML_DEGREE,
        },
      ],
      connectionSites: [
        { angle: 0, position: { x: 0, y: 0 } },
        { angle: 0, position: { x: 0, y: 0 } },
      ],
      textRectangle: { left: 50_000, top: 1_000, right: 200_000, bottom: 75_000 },
      paths: [
        {
          width: 100_000,
          height: 100_000,
          fill: 'norm',
          stroke: false,
          extrusionOk: true,
          commands: [
            { kind: 'moveTo', point: { x: 50_000, y: 0 } },
            { kind: 'lineTo', point: { x: 200_000, y: 100_000 } },
            {
              kind: 'quadraticBezierTo',
              control: { x: 100_000, y: 0 },
              end: { x: 200_000, y: 50_000 },
            },
            {
              kind: 'cubicBezierTo',
              control1: { x: 1, y: 2 },
              control2: { x: 50_000, y: 75_000 },
              end: { x: 200_000, y: 100_000 },
            },
            {
              kind: 'arcTo',
              widthRadius: 10_000,
              heightRadius: 50_000,
              startAngle: 90 * OOXML_DEGREE,
              sweepAngle: 180 * OOXML_DEGREE,
            },
            { kind: 'close' },
          ],
        },
        { width: 1, height: 1, commands: [] },
      ],
    });
    expect(evaluated.paths[0]?.commands[1]).toEqual({
      kind: 'lineTo',
      point: { x: 200_000, y: 100_000 },
    });
  });

  it('materializes the default text rectangle and preserves omitted lists', () => {
    const evaluated = evaluateCustomGeometry(
      { paths: [{ width: 101, height: 67, commands: [] }] },
      { width: 200_000, height: 100_000 },
    );

    expect(evaluated.textRectangle).toEqual({
      left: 0,
      top: 0,
      right: 200_000,
      bottom: 100_000,
    });
    expect(Object.hasOwn(evaluated, 'adjustments')).toBe(false);
    expect(Object.hasOwn(evaluated, 'guides')).toBe(false);
    expect(Object.hasOwn(evaluated, 'handles')).toBe(false);
    expect(Object.hasOwn(evaluated, 'connectionSites')).toBe(false);
  });

  it('detaches inputs and recursively freezes every output branch', () => {
    const source = structuredClone({
      guides: [{ name: 'g1', formula: { operator: 'val' as const, operands: [10] as [number] } }],
      handles: [{ kind: 'xy' as const, position: { x: 'g1', y: 0 }, minX: 0 }],
      connectionSites: [{ angle: 0, position: { x: 'g1', y: 0 } }],
      textRectangle: { left: 0, top: 0, right: 'g1', bottom: 'b' },
      paths: [{
        width: 10,
        height: 10,
        commands: [{ kind: 'moveTo' as const, point: { x: 'g1', y: 0 } }],
      }],
    });
    const context = { width: 100, height: 50 };
    const evaluated = evaluateCustomGeometry(source, context);

    source.guides[0]!.name = 'changed';
    source.handles[0]!.position.x = 'changed';
    source.paths[0]!.commands[0]!.point.x = 'changed';
    context.width = 999;

    expect(evaluated.context.width).toBe(100);
    expect(evaluated.guides?.[0]).toEqual({ name: 'g1', value: 10 });
    expect(evaluated.handles?.[0]?.position.x).toBe(10);
    expect(evaluated.paths[0]?.commands[0]).toEqual({
      kind: 'moveTo',
      point: { x: 10, y: 0 },
    });
    expectDeepFrozen(evaluated);
  });

  it('strictly validates evaluation context and geometry input', () => {
    const geometry: CustomGeometry = { paths: [{ width: 1, height: 1, commands: [] }] };
    const inherited = Object.create({ width: 1, height: 1 }) as CustomGeometryEvaluationContext;
    const accessor = Object.defineProperty({ height: 1 }, 'width', { get: () => 1 });
    const symbol = { width: 1, height: 1, [Symbol('unsafe')]: true };
    const invalidContexts: unknown[] = [
      null,
      [],
      inherited,
      accessor,
      symbol,
      { width: 1 },
      { width: 1, height: 1, extra: true },
      { width: '1', height: 1 },
      { width: Number.POSITIVE_INFINITY, height: 1 },
      { width: 1.5, height: 1 },
      { width: Number.MAX_SAFE_INTEGER + 1, height: 1 },
      { width: 0, height: 1 },
      { width: 1, height: -1 },
    ];

    for (const context of invalidContexts) {
      expect(() => evaluateCustomGeometry(
        geometry,
        context as CustomGeometryEvaluationContext,
      ), JSON.stringify(context)).toThrow();
    }
    expect(() => evaluateCustomGeometry(
      { paths: [{ width: 1, height: 1, commands: [] }], unknown: true } as never,
      { width: 1, height: 1 },
    )).toThrow(TypeError);
  });

  it('rejects unknown tokens and non-positive evaluated arc radii', () => {
    const unknown = (): unknown => evaluateCustomGeometry(
      {
        paths: [{
          width: 1,
          height: 1,
          commands: [{ kind: 'moveTo', point: { x: 'missing', y: 0 } }],
        }],
      },
      { width: 100, height: 100 },
    );
    const invalidRadius = (radius: number): (() => unknown) => () => evaluateCustomGeometry(
      {
        guides: [{ name: 'rad', formula: { operator: 'val', operands: [radius] } }],
        paths: [{
          width: 1,
          height: 1,
          commands: [
            { kind: 'moveTo', point: { x: 0, y: 0 } },
            {
              kind: 'arcTo',
              widthRadius: 'rad',
              heightRadius: 1,
              startAngle: 0,
              sweepAngle: 1,
            },
          ],
        }],
      },
      { width: 100, height: 100 },
    );

    expectEvaluationError(unknown, 'unknown-token', undefined, 'missing');
    expectEvaluationError(invalidRadius(0), 'invalid-domain', undefined, 'rad');
    expectEvaluationError(invalidRadius(-1), 'invalid-domain', undefined, 'rad');
  });

  it('rejects non-finite results propagated through public guide evaluation', () => {
    const square = (
      name: string,
      operand: number | string,
    ): NonNullable<CustomGeometry['guides']>[number] => ({
      name,
      formula: { operator: '*/', operands: [operand, operand, 1] },
    });
    const operation = (): unknown => evaluateCustomGeometry(
      {
        guides: [
          square('g1', Number.MAX_SAFE_INTEGER),
          square('g2', 'g1'),
          square('g3', 'g2'),
          square('g4', 'g3'),
          square('g5', 'g4'),
        ],
        paths: [{ width: 1, height: 1, commands: [] }],
      },
      { width: 100, height: 100 },
    );

    expectEvaluationError(operation, 'non-finite-result', 'g5');
  });
});

function expectEvaluationError(
  operation: () => unknown,
  code: CustomGeometryEvaluationError['code'],
  guideName?: string,
  token?: string,
): void {
  try {
    operation();
    throw new Error('Expected custom geometry evaluation to fail');
  } catch (error) {
    expect(error).toBeInstanceOf(CustomGeometryEvaluationError);
    expect(error).toMatchObject({ code, guideName, token });
  }
}

function geometryWithGuides(
  adjustments: CustomGeometry['adjustments'],
  guides: CustomGeometry['guides'],
): CustomGeometry {
  return {
    ...(adjustments ? { adjustments } : {}),
    ...(guides ? { guides } : {}),
    paths: [],
  };
}

function expectDeepFrozen(value: unknown, seen = new Set<unknown>()): void {
  if (!value || (typeof value !== 'object' && typeof value !== 'function') || seen.has(value)) {
    return;
  }
  seen.add(value);
  expect(Object.isFrozen(value)).toBe(true);
  for (const nested of Object.values(value)) expectDeepFrozen(nested, seen);
}
