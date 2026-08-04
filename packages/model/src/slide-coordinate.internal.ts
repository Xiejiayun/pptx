import type { Emu, SlideSize } from './units.js';

export type SlideCoordinateAxis = 'horizontal' | 'vertical';

const PERCENTAGE = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?%$/iu;

export function resolveSlideCoordinate(
  value: unknown,
  axis: SlideCoordinateAxis,
  slideSize: Readonly<SlideSize> | undefined,
  fallback: Emu,
  context: string,
): Emu {
  if (value === undefined) return fallback;
  if (typeof value === 'number') return normalizeEmu(value, context);
  if (typeof value !== 'string' || !PERCENTAGE.test(value)) {
    throw new TypeError(`${context} must be an absolute EMU or percentage`);
  }
  if (slideSize === undefined) {
    throw new TypeError(`${context} percentage requires a slide size`);
  }
  const percentage = Number(value.slice(0, -1));
  if (!Number.isFinite(percentage)) {
    throw new TypeError(`${context} percentage must be finite`);
  }
  const dimension = axis === 'horizontal' ? slideSize.width : slideSize.height;
  return normalizeEmu(percentage * dimension / 100, context);
}

function normalizeEmu(value: number, context: string): Emu {
  if (!Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  const rounded = Math.round(value);
  if (!Number.isSafeInteger(rounded)) {
    throw new RangeError(`${context} must resolve to a safe EMU integer`);
  }
  return (Object.is(rounded, -0) ? 0 : rounded) as Emu;
}
