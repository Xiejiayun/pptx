import { describe, expect, it } from 'vitest';
import {
  CustomGeometryEvaluationError,
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
} from './custom-geometry-evaluator.internal.js';

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

function expectEvaluationError(
  operation: () => unknown,
  code: CustomGeometryEvaluationError['code'],
  guideName: string,
): void {
  try {
    operation();
    throw new Error('Expected custom geometry evaluation to fail');
  } catch (error) {
    expect(error).toBeInstanceOf(CustomGeometryEvaluationError);
    expect(error).toMatchObject({ code, guideName });
  }
}
