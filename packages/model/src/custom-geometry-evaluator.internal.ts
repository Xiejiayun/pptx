import {
  CustomGeometryEvaluationError,
  type CustomGeometry,
  type CustomGeometryEvaluationContext,
  type CustomGeometryFormula,
  type CustomGeometryGuide,
  type CustomGeometryValue,
  type EvaluatedCustomGeometryGuide,
} from './custom-geometry.js';

const OOXML_DEGREE = 60_000;

export function evaluateBuiltInGuide(
  token: string,
  context: Readonly<CustomGeometryEvaluationContext>,
): number | undefined {
  const { width, height } = context;
  const shortSide = Math.min(width, height);
  switch (token) {
    case '3cd4': return 270 * OOXML_DEGREE;
    case '3cd8': return 135 * OOXML_DEGREE;
    case '5cd8': return 225 * OOXML_DEGREE;
    case '7cd8': return 315 * OOXML_DEGREE;
    case 'b':
    case 'h': return height;
    case 'cd2': return 180 * OOXML_DEGREE;
    case 'cd4': return 90 * OOXML_DEGREE;
    case 'cd8': return 45 * OOXML_DEGREE;
    case 'hc': return width / 2;
    case 'hd2': return height / 2;
    case 'hd3': return height / 3;
    case 'hd4': return height / 4;
    case 'hd5': return height / 5;
    case 'hd6': return height / 6;
    case 'hd8': return height / 8;
    case 'l':
    case 't': return 0;
    case 'ls': return Math.max(width, height);
    case 'r':
    case 'w': return width;
    case 'ss': return shortSide;
    case 'ssd2': return shortSide / 2;
    case 'ssd4': return shortSide / 4;
    case 'ssd6': return shortSide / 6;
    case 'ssd8': return shortSide / 8;
    case 'ssd16': return shortSide / 16;
    case 'ssd32': return shortSide / 32;
    case 'vc': return height / 2;
    case 'wd2': return width / 2;
    case 'wd3': return width / 3;
    case 'wd4': return width / 4;
    case 'wd5': return width / 5;
    case 'wd6': return width / 6;
    case 'wd8': return width / 8;
    case 'wd10': return width / 10;
    case 'wd32': return width / 32;
    default: return undefined;
  }
}

export function evaluateFormulaValue(
  formula: Readonly<CustomGeometryFormula>,
  resolve: (value: CustomGeometryValue, location: string) => number,
  guideName: string,
): number {
  const operands = formula.operands.map((value, index) => resolve(
    value,
    `Custom geometry guide ${guideName} operand ${index}`,
  ));
  const x = operands[0]!;
  const y = operands[1] ?? 0;
  const z = operands[2] ?? 0;
  let result: number;
  switch (formula.operator) {
    case 'val': result = x; break;
    case 'abs': result = Math.abs(x); break;
    case 'sqrt':
      if (x < 0) {
        throw new CustomGeometryEvaluationError(
          'invalid-domain',
          `Custom geometry guide ${guideName} cannot take the square root of ${x}`,
          guideName,
        );
      }
      result = Math.sqrt(x);
      break;
    case 'at2': result = radiansToOoxmlAngle(Math.atan2(y, x)); break;
    case 'cos': result = x * Math.cos(ooxmlAngleToRadians(y)); break;
    case 'max': result = Math.max(x, y); break;
    case 'min': result = Math.min(x, y); break;
    case 'sin': result = x * Math.sin(ooxmlAngleToRadians(y)); break;
    case 'tan': result = x * Math.tan(ooxmlAngleToRadians(y)); break;
    case '*/': result = z === 0 ? 0 : x * y / z; break;
    case '+-': result = x + y - z; break;
    case '+/': result = z === 0 ? 0 : (x + y) / z; break;
    case '?:': result = x > 0 ? y : z; break;
    case 'cat2': result = x * Math.cos(Math.atan2(z, y)); break;
    case 'mod': result = Math.sqrt(x * x + y * y + z * z); break;
    case 'pin': result = Math.max(x, Math.min(y, z)); break;
    case 'sat2': result = x * Math.sin(Math.atan2(z, y)); break;
  }
  return normalizeFormulaResult(result, guideName);
}

export interface EvaluatedGuideEnvironment {
  readonly adjustments?: readonly EvaluatedCustomGeometryGuide[];
  readonly guides?: readonly EvaluatedCustomGeometryGuide[];
  resolve(value: CustomGeometryValue, location: string): number;
}

export function evaluateGuideEnvironment(
  geometry: Readonly<CustomGeometry>,
  context: Readonly<CustomGeometryEvaluationContext>,
): EvaluatedGuideEnvironment {
  const adjustmentGuides = geometry.adjustments ?? [];
  const shapeGuides = geometry.guides ?? [];
  const guides = [...adjustmentGuides, ...shapeGuides];
  auditGuideDependencies(guides, context);

  const values = new Map<string, number>();
  const resolve = (value: CustomGeometryValue, location: string): number => {
    if (typeof value === 'number') return Object.is(value, -0) ? 0 : value;
    if (values.has(value)) return values.get(value)!;
    const builtIn = evaluateBuiltInGuide(value, context);
    if (builtIn !== undefined) return builtIn;
    throw new CustomGeometryEvaluationError(
      'unknown-token',
      `${location} references unknown token ${value}`,
      undefined,
      value,
    );
  };
  const evaluateGuides = (
    source: readonly Readonly<CustomGeometryGuide>[],
  ): readonly EvaluatedCustomGeometryGuide[] | undefined => {
    if (source.length === 0) return undefined;
    return Object.freeze(source.map((guide) => {
      const value = evaluateFormulaValue(guide.formula, resolve, guide.name);
      values.set(guide.name, value);
      return Object.freeze({ name: guide.name, value });
    }));
  };
  const adjustments = evaluateGuides(adjustmentGuides);
  const evaluatedGuides = evaluateGuides(shapeGuides);
  return Object.freeze({
    ...(adjustments ? { adjustments } : {}),
    ...(evaluatedGuides ? { guides: evaluatedGuides } : {}),
    resolve,
  });
}

interface GuideDependency {
  readonly target: number;
  readonly token: string;
}

interface InvalidGuideReference {
  readonly source: number;
  readonly token: string;
}

function auditGuideDependencies(
  guides: readonly Readonly<CustomGeometryGuide>[],
  context: Readonly<CustomGeometryEvaluationContext>,
): void {
  const indexes = new Map(guides.map(({ name }, index) => [name, index]));
  const edges: GuideDependency[][] = guides.map(() => []);
  const forwardReferences: InvalidGuideReference[] = [];
  const unknownTokens: InvalidGuideReference[] = [];
  for (const [source, guide] of guides.entries()) {
    for (const operand of guide.formula.operands) {
      if (typeof operand !== 'string') continue;
      const target = indexes.get(operand);
      const isBuiltIn = evaluateBuiltInGuide(operand, context) !== undefined;
      if (target === undefined) {
        if (!isBuiltIn) unknownTokens.push({ source, token: operand });
        continue;
      }
      if (target < source) {
        edges[source]!.push({ target, token: operand });
        continue;
      }
      if (isBuiltIn) continue;
      edges[source]!.push({ target, token: operand });
      if (target > source) forwardReferences.push({ source, token: operand });
    }
  }

  const cycle = findDependencyCycle(edges);
  if (cycle) {
    const source = Math.min(...cycle);
    const dependency = edges[source]!.find(({ target }) => cycle.has(target))!;
    const guideName = guides[source]!.name;
    throw new CustomGeometryEvaluationError(
      'cyclic-reference',
      `Custom geometry guide ${guideName} participates in a dependency cycle through ${dependency.token}`,
      guideName,
      dependency.token,
    );
  }
  const forward = forwardReferences[0];
  if (forward) {
    const guideName = guides[forward.source]!.name;
    throw new CustomGeometryEvaluationError(
      'forward-reference',
      `Custom geometry guide ${guideName} references later guide ${forward.token}`,
      guideName,
      forward.token,
    );
  }
  const unknown = unknownTokens[0];
  if (unknown) {
    const guideName = guides[unknown.source]!.name;
    throw new CustomGeometryEvaluationError(
      'unknown-token',
      `Custom geometry guide ${guideName} references unknown token ${unknown.token}`,
      guideName,
      unknown.token,
    );
  }
}

function findDependencyCycle(
  edges: readonly (readonly GuideDependency[])[],
): ReadonlySet<number> | undefined {
  const states = new Uint8Array(edges.length);
  const stack: number[] = [];
  let cycle: ReadonlySet<number> | undefined;
  const visit = (source: number): void => {
    states[source] = 1;
    stack.push(source);
    for (const { target } of edges[source]!) {
      if (states[target] === 0) visit(target);
      else if (states[target] === 1 && !cycle) {
        cycle = new Set(stack.slice(stack.lastIndexOf(target)));
      }
      if (cycle) return;
    }
    stack.pop();
    states[source] = 2;
  };
  for (let index = 0; index < edges.length && !cycle; index += 1) {
    if (states[index] === 0) visit(index);
  }
  return cycle;
}

function ooxmlAngleToRadians(value: number): number {
  return value / OOXML_DEGREE * Math.PI / 180;
}

function radiansToOoxmlAngle(value: number): number {
  return value * 180 / Math.PI * OOXML_DEGREE;
}

function normalizeFormulaResult(value: number, guideName: string): number {
  if (!Number.isFinite(value)) {
    throw new CustomGeometryEvaluationError(
      'non-finite-result',
      `Custom geometry guide ${guideName} produced a non-finite result`,
      guideName,
    );
  }
  return Object.is(value, -0) ? 0 : value;
}
