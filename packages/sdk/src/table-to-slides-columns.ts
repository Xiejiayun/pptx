import { inches } from '@pptx/model';
import type { HtmlTableCellSnapshot, HtmlTableSnapshot } from './table-to-slides.js';

const MAX_PHYSICAL_COLUMNS = 1_000_000;
const HTML_INCH_NUMBER = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i;

export interface ResolvedHtmlTableColumns {
  readonly widths: readonly number[];
  readonly width: number;
}

export function resolveHtmlTableColumns(
  snapshot: Readonly<HtmlTableSnapshot>,
  targetWidth: number,
  explicit?: number | readonly number[],
): Readonly<ResolvedHtmlTableColumns> {
  const width = requirePositiveSafeInteger(targetWidth, 'HTML table target width');
  const sourceRow = snapshot.rows[snapshot.widthSourceRowIndex];
  if (!sourceRow || sourceRow.length === 0) {
    throw new RangeError('HTML table width source row must contain cells');
  }
  const pixelWeights = expandPixelWeights(sourceRow);
  const columnCount = pixelWeights.length;
  if (width < columnCount) {
    throw new RangeError('HTML table target width must provide at least one EMU per column');
  }
  if (explicit !== undefined) return resolveExplicit(columnCount, width, explicit);

  const { fixed, minimum } = readHeaderConstraints(snapshot, columnCount);
  const flexibleIndexes = Array.from({ length: columnCount }, (_, index) => index)
    .filter((index) => fixed[index] === undefined);
  if (flexibleIndexes.length === 0) {
    const fixedWidths = fixed as number[];
    const fixedWidth = safeSum(fixedWidths, 'HTML table fixed column widths');
    return freezeResult(fixedWidths, fixedWidth);
  }
  if (pixelWeights.every((weight) => weight === 0)) {
    throw new RangeError(
      'HTML table is hidden or has zero layout width; provide explicit columnWidths',
    );
  }

  const fixedSum = safeSum(
    fixed.filter((value): value is number => value !== undefined),
    'HTML table fixed column widths',
  );
  const flexibleMinimums = flexibleIndexes.map((index) => Math.max(1, minimum[index] ?? 0));
  const minimumSum = safeSum(flexibleMinimums, 'HTML table minimum column widths');
  const requiredWidth = safeAdd(fixedSum, minimumSum, 'HTML table constrained width');
  const actualWidth = Math.max(width, requiredWidth);
  if (!Number.isSafeInteger(actualWidth)) {
    throw new RangeError('HTML table constrained width must be a safe integer');
  }
  const flexibleWidth = actualWidth - fixedSum;
  const flexibleWeights = flexibleIndexes.map((index) => pixelWeights[index]!);
  const flexibleWidths = allocateWithMinimums(
    flexibleWidth,
    flexibleWeights,
    flexibleMinimums,
  );
  const widths = Array.from({ length: columnCount }, (_, index) => fixed[index] ?? 0);
  for (let index = 0; index < flexibleIndexes.length; index += 1) {
    widths[flexibleIndexes[index]!] = flexibleWidths[index]!;
  }
  assertResolvedWidths(widths, actualWidth);
  return freezeResult(widths, actualWidth);
}

function expandPixelWeights(
  row: readonly Readonly<HtmlTableCellSnapshot>[],
): number[] {
  const result: number[] = [];
  for (const cell of row) {
    const span = cell.colspan ?? 1;
    if (!Number.isSafeInteger(span) || span < 1) {
      throw new TypeError('HTML table cell colspan must be a positive safe integer');
    }
    if (result.length + span > MAX_PHYSICAL_COLUMNS) {
      throw new RangeError(
        `HTML table cannot contain more than ${MAX_PHYSICAL_COLUMNS} physical columns`,
      );
    }
    if (
      typeof cell.offsetWidth !== 'number'
      || !Number.isFinite(cell.offsetWidth)
      || cell.offsetWidth < 0
    ) {
      throw new TypeError('HTML table cell offsetWidth must be a non-negative finite number');
    }
    const weight = cell.offsetWidth / span;
    for (let index = 0; index < span; index += 1) result.push(weight);
  }
  return result;
}

function resolveExplicit(
  columnCount: number,
  targetWidth: number,
  explicit: number | readonly number[],
): Readonly<ResolvedHtmlTableColumns> {
  const values = Array.isArray(explicit)
    ? readDensePlainArray(explicit, 'HTML table explicit columnWidths')
    : Array.from({ length: columnCount }, () => explicit);
  if (values.length !== columnCount) {
    throw new RangeError('HTML table explicit columnWidths must match the physical column count');
  }
  const widths = values.map((value, index) => requirePositiveSafeInteger(
    value,
    `HTML table explicit columnWidths ${index}`,
  ));
  const sum = safeSum(widths, 'HTML table explicit columnWidths');
  if (sum !== targetWidth) {
    throw new RangeError('HTML table explicit columnWidths must equal the target width');
  }
  return freezeResult(widths, targetWidth);
}

function readHeaderConstraints(
  snapshot: Readonly<HtmlTableSnapshot>,
  columnCount: number,
): Readonly<{
  fixed: readonly (number | undefined)[];
  minimum: readonly (number | undefined)[];
}> {
  const fixed: (number | undefined)[] = Array.from({ length: columnCount });
  const minimum: (number | undefined)[] = Array.from({ length: columnCount });
  const headerRow = snapshot.rows.find((row) => row.some((cell) => cell.header));
  if (!headerRow) return Object.freeze({ fixed, minimum });
  let columnIndex = 0;
  for (const cell of headerRow) {
    const span = cell.colspan ?? 1;
    if (columnIndex + span > columnCount) {
      throw new RangeError('HTML table header has more physical columns than the width source row');
    }
    if (cell.header && cell.pptxWidth !== undefined) {
      const total = parseWidthAttribute(cell.pptxWidth, 'data-pptx-width', false);
      const values = distributeConstraint(total, span, 'data-pptx-width', true);
      for (let offset = 0; offset < span; offset += 1) {
        fixed[columnIndex + offset] = values[offset]!;
      }
    } else if (cell.header && cell.pptxMinWidth !== undefined) {
      const total = parseWidthAttribute(cell.pptxMinWidth, 'data-pptx-min-width', true);
      const values = distributeConstraint(total, span, 'data-pptx-min-width', false);
      for (let offset = 0; offset < span; offset += 1) {
        minimum[columnIndex + offset] = values[offset]!;
      }
    }
    columnIndex += span;
  }
  return Object.freeze({ fixed, minimum });
}

function parseWidthAttribute(
  value: string,
  name: string,
  allowZero: boolean,
): number {
  const trimmed = value.trim();
  if (trimmed.length === 0) throw new TypeError(`HTML table ${name} cannot be empty`);
  if (!HTML_INCH_NUMBER.test(trimmed)) {
    throw new TypeError(`HTML table ${name} must be a decimal number`);
  }
  const numeric = Number(trimmed);
  if (!Number.isFinite(numeric)) throw new TypeError(`HTML table ${name} must be finite`);
  if (allowZero ? numeric < 0 : numeric <= 0) {
    throw new RangeError(
      `HTML table ${name} must be ${allowZero ? 'non-negative' : 'positive'}`,
    );
  }
  const emu = inches(numeric);
  if (!Number.isSafeInteger(emu) || emu < 0) {
    throw new RangeError(`HTML table ${name} must convert to a safe EMU value`);
  }
  if (!allowZero && emu === 0) {
    throw new RangeError(`HTML table ${name} must be at least one EMU`);
  }
  return emu;
}

function distributeConstraint(
  total: number,
  span: number,
  context: string,
  requirePositive: boolean,
): readonly number[] {
  if (requirePositive && total < span) {
    throw new RangeError(`HTML table ${context} must provide at least one EMU per spanned column`);
  }
  const quotient = Math.floor(total / span);
  const remainder = total - quotient * span;
  return Object.freeze(Array.from(
    { length: span },
    (_, index) => quotient + (index < remainder ? 1 : 0),
  ));
}

function allocateWithMinimums(
  total: number,
  weights: readonly number[],
  minimums: readonly number[],
): readonly number[] {
  const result = Array.from({ length: weights.length }, () => 0);
  let active = Array.from({ length: weights.length }, (_, index) => index);
  let remaining = total;
  while (active.length > 0) {
    const activeWeights = active.map((index) => weights[index]!);
    const effectiveWeights = activeWeights.some((weight) => weight > 0)
      ? activeWeights
      : activeWeights.map(() => 1);
    const allocation = allocateLargestRemainder(remaining, effectiveWeights);
    const belowMinimum = active.filter(
      (_, activeIndex) => allocation[activeIndex]! < minimums[active[activeIndex]!]!,
    );
    if (belowMinimum.length === 0) {
      for (let index = 0; index < active.length; index += 1) {
        result[active[index]!] = allocation[index]!;
      }
      break;
    }
    const below = new Set(belowMinimum);
    for (const index of belowMinimum) {
      const minimum = minimums[index]!;
      result[index] = minimum;
      remaining -= minimum;
    }
    active = active.filter((index) => !below.has(index));
  }
  if (remaining < 0) throw new RangeError('HTML table minimum widths exceed available width');
  const sum = safeSum(result, 'HTML table flexible column widths');
  if (sum !== total) throw new RangeError('HTML table flexible column allocation is inexact');
  return Object.freeze(result);
}

function allocateLargestRemainder(
  total: number,
  weights: readonly number[],
): number[] {
  if (weights.length === 0) return [];
  let weightSum = 0;
  for (const weight of weights) {
    if (!Number.isFinite(weight) || weight < 0) {
      throw new TypeError('HTML table pixel weights must be non-negative and finite');
    }
    weightSum += weight;
  }
  if (!Number.isFinite(weightSum) || weightSum <= 0) {
    throw new RangeError('HTML table pixel weights must contain a positive value');
  }
  const raw = weights.map((weight) => total * (weight / weightSum));
  const result = raw.map(Math.floor);
  let delta = total - result.reduce((sum, value) => sum + value, 0);
  const descending = raw.map((value, index) => ({
    index,
    fraction: value - Math.floor(value),
  })).sort((left, right) => right.fraction - left.fraction || left.index - right.index);
  if (delta > 0) {
    const quotient = Math.floor(delta / descending.length);
    if (quotient > 0) {
      for (const { index } of descending) result[index] = result[index]! + quotient;
      delta -= quotient * descending.length;
    }
    for (let index = 0; index < delta; index += 1) {
      const resultIndex = descending[index]!.index;
      result[resultIndex] = result[resultIndex]! + 1;
    }
  } else if (delta < 0) {
    const ascending = [...descending].reverse();
    while (delta < 0) {
      let changed = false;
      for (const { index } of ascending) {
        if (result[index]! === 0) continue;
        result[index] = result[index]! - 1;
        delta += 1;
        changed = true;
        if (delta === 0) break;
      }
      if (!changed) throw new RangeError('HTML table proportional allocation is inexact');
    }
  }
  return result;
}

function readDensePlainArray(value: readonly number[], context: string): readonly unknown[] {
  if (Object.getPrototypeOf(value) !== Array.prototype) {
    throw new TypeError(`${context} must be a plain array`);
  }
  const allowed = new Set([
    'length',
    ...Array.from({ length: value.length }, (_, index) => String(index)),
  ]);
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !allowed.has(key)) {
      throw new TypeError(`${context} contains unsupported property ${String(key)}`);
    }
  }
  return Array.from({ length: value.length }, (_, index) => {
    const descriptor = Object.getOwnPropertyDescriptor(value, String(index));
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`${context} must be dense and contain only data items`);
    }
    return descriptor.value;
  });
}

function requirePositiveSafeInteger(value: unknown, context: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  if (!Number.isSafeInteger(value)) throw new RangeError(`${context} must be a safe integer`);
  if (value <= 0) throw new RangeError(`${context} must be positive`);
  return value;
}

function safeSum(values: readonly number[], context: string): number {
  return values.reduce((sum, value) => safeAdd(sum, value, context), 0);
}

function safeAdd(left: number, right: number, context: string): number {
  const sum = left + right;
  if (!Number.isSafeInteger(sum)) throw new RangeError(`${context} exceeds the safe integer range`);
  return sum;
}

function assertResolvedWidths(widths: readonly number[], width: number): void {
  for (const value of widths) {
    if (!Number.isSafeInteger(value) || value <= 0) {
      throw new RangeError('HTML table resolved widths must be positive safe integers');
    }
  }
  if (safeSum(widths, 'HTML table resolved widths') !== width) {
    throw new RangeError('HTML table resolved widths must exactly equal the table width');
  }
}

function freezeResult(
  widths: readonly number[],
  width: number,
): Readonly<ResolvedHtmlTableColumns> {
  assertResolvedWidths(widths, width);
  return Object.freeze({ widths: Object.freeze([...widths]), width });
}
