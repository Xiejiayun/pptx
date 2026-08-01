const EDGE_KEYS = ['left', 'top', 'right', 'bottom'] as const;
const EDGE_KEY_SET = new Set<string>(EDGE_KEYS);
const PERCENT_SCALE = 1_000;
const FULL_PERCENT = 100 * PERCENT_SCALE;
const INT32_MIN = -2_147_483_648;
const INT32_MAX = 2_147_483_647;

export interface NormalizedImageSourceRectangle {
  readonly left: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
}

export function normalizeImageSourceRectangle(
  value: unknown,
  context: string,
): Readonly<NormalizedImageSourceRectangle> {
  const values = readRectangle(value, context);
  const raw = {
    left: normalizeEdge(values.left, `${context} left`),
    top: normalizeEdge(values.top, `${context} top`),
    right: normalizeEdge(values.right, `${context} right`),
    bottom: normalizeEdge(values.bottom, `${context} bottom`),
  };
  if (raw.left + raw.right >= FULL_PERCENT) {
    throw new RangeError(`${context} left and right must leave a positive source width`);
  }
  if (raw.top + raw.bottom >= FULL_PERCENT) {
    throw new RangeError(`${context} top and bottom must leave a positive source height`);
  }
  return Object.freeze({
    left: percentage(raw.left),
    top: percentage(raw.top),
    right: percentage(raw.right),
    bottom: percentage(raw.bottom),
  });
}

export function renderImageSourceRectangle(
  value: Readonly<NormalizedImageSourceRectangle>,
  prefix = 'a',
): string {
  const name = prefix.length === 0 ? 'srcRect' : `${prefix}:srcRect`;
  return `<${name} l="${rawPercentage(value.left)}" t="${rawPercentage(value.top)}" `
    + `r="${rawPercentage(value.right)}" b="${rawPercentage(value.bottom)}"/>`;
}

function readRectangle(value: unknown, context: string): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${context} must be an object`);
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError(`${context} must be an ordinary object`);
  }
  const result = Object.create(null) as Record<string, unknown>;
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !EDGE_KEY_SET.has(key)) {
      throw new TypeError(`${context} contains unsupported property ${String(key)}`);
    }
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`${context} ${key} must be a data property`);
    }
    result[key] = descriptor.value;
  }
  for (const key of EDGE_KEYS) {
    if (!Object.hasOwn(result, key)) throw new TypeError(`${context} ${key} is required`);
  }
  return result;
}

function normalizeEdge(value: unknown, context: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  const raw = Math.round(value * PERCENT_SCALE);
  if (!Number.isSafeInteger(raw) || raw < INT32_MIN || raw > INT32_MAX) {
    throw new RangeError(`${context} must fit the OOXML Int32 percentage range`);
  }
  if (raw >= FULL_PERCENT) throw new RangeError(`${context} must be less than 100 percent`);
  return raw === 0 ? 0 : raw;
}

function percentage(raw: number): number {
  const value = raw / PERCENT_SCALE;
  return value === 0 ? 0 : value;
}

function rawPercentage(value: number): number {
  const raw = Math.round(value * PERCENT_SCALE);
  return raw === 0 ? 0 : raw;
}
