import type {
  CustomGeometry,
  CustomGeometryEvaluationContext,
  EvaluatedCustomGeometry,
} from './custom-geometry.js';
import { evaluateCustomGeometryTree } from './custom-geometry-evaluator.internal.js';
import { normalizeCustomGeometry } from './custom-geometry.internal.js';

const CONTEXT_KEYS = new Set(['width', 'height']);

export function evaluateCustomGeometry(
  geometry: CustomGeometry,
  context: CustomGeometryEvaluationContext,
): EvaluatedCustomGeometry {
  const normalizedGeometry = normalizeCustomGeometry(geometry, 'Custom geometry evaluation');
  const normalizedContext = normalizeEvaluationContext(context);
  return evaluateCustomGeometryTree(normalizedGeometry, normalizedContext);
}

function normalizeEvaluationContext(value: unknown): Readonly<CustomGeometryEvaluationContext> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError('Custom geometry evaluation context must be an object');
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError('Custom geometry evaluation context must be an ordinary object');
  }
  const result = Object.create(null) as Record<string, unknown>;
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !CONTEXT_KEYS.has(key)) {
      throw new TypeError(
        `Custom geometry evaluation context contains unsupported property ${String(key)}`,
      );
    }
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(
        `Custom geometry evaluation context property ${key} must be a data property`,
      );
    }
    result[key] = descriptor.value;
  }
  for (const key of CONTEXT_KEYS) {
    if (!Object.hasOwn(result, key)) {
      throw new TypeError(`Custom geometry evaluation context is missing property ${key}`);
    }
  }
  return Object.freeze({
    width: readPositiveSafeInteger(result.width, 'width'),
    height: readPositiveSafeInteger(result.height, 'height'),
  });
}

function readPositiveSafeInteger(value: unknown, property: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`Custom geometry evaluation context ${property} must be finite`);
  }
  if (!Number.isSafeInteger(value)) {
    throw new RangeError(`Custom geometry evaluation context ${property} must be a safe integer`);
  }
  if (value <= 0) {
    throw new RangeError(`Custom geometry evaluation context ${property} must be positive`);
  }
  return value;
}
