import type {
  NormalizedTableCell,
  NormalizedTableDefinition,
} from './table-create.internal.js';
import { EMU_PER_INCH, type SlideSize } from './units.js';

const DEFAULT_PAGE_MARGIN = EMU_PER_INCH / 2;
const AUTO_PAGE_CONTROL_KEYS = [
  'autoPageRepeatHeader',
  'autoPageHeaderRows',
  'autoPageSlideStartY',
  'slideMargin',
  'autoPageCharWeight',
  'autoPageLineWeight',
] as const;

export interface NormalizedTableAutoPageRequest {
  readonly repeatHeader: boolean;
  readonly headerRows: number;
  readonly slideStartY?: number;
  readonly slideMargin?: readonly [number, number, number, number];
  readonly charWeight?: number;
  readonly lineWeight?: number;
  readonly measureContent: boolean;
}

export interface TableAutoPageMargins {
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
  readonly left: number;
}

export interface TableAutoPageLayoutRegion {
  readonly firstY: number;
  readonly continuationY: number;
  readonly bottomEdge: number;
  readonly firstCapacity: number;
  readonly continuationCapacity: number;
}

interface NormalizeTableAutoPageContext {
  readonly rowCount: number;
  readonly rowHeights: readonly number[];
  readonly autoRowHeight: boolean;
  readonly hasPlaceholder: boolean;
  readonly hasCellMeasurementWeights: boolean;
}

interface TableRowBlock {
  readonly start: number;
  readonly end: number;
  readonly height: number;
}

export function normalizeTableAutoPageRequest(
  options: Readonly<Record<string, unknown>>,
  context: Readonly<NormalizeTableAutoPageContext>,
): Readonly<NormalizedTableAutoPageRequest> | undefined {
  const enabled = normalizeTableAutoPageEnabled(options.autoPage);
  const hasControls = AUTO_PAGE_CONTROL_KEYS.some((key) => options[key] !== undefined);
  if (!enabled) {
    if (hasControls || context.hasCellMeasurementWeights) {
      throw new TypeError('Table auto-page controls require autoPage to be true');
    }
    return undefined;
  }
  if (context.hasPlaceholder) {
    throw new TypeError('Table autoPage does not support placeholder placement');
  }
  const repeatHeaderValue = options.autoPageRepeatHeader;
  if (repeatHeaderValue !== undefined && typeof repeatHeaderValue !== 'boolean') {
    throw new TypeError('Table autoPageRepeatHeader must be a boolean');
  }
  const repeatHeader = repeatHeaderValue ?? false;
  const requestedHeaderRows = options.autoPageHeaderRows;
  if (!repeatHeader && requestedHeaderRows !== undefined) {
    throw new TypeError(
      'Table autoPageHeaderRows requires autoPageRepeatHeader to be true',
    );
  }
  const headerRows = repeatHeader
    ? normalizeHeaderRows(requestedHeaderRows, context.rowCount)
    : 0;
  const slideStartY = options.autoPageSlideStartY === undefined
    ? undefined
    : normalizeNonNegativeSafeInteger(
        options.autoPageSlideStartY,
        'Table autoPageSlideStartY',
      );
  const slideMargin = options.slideMargin === undefined
    ? undefined
    : normalizeSlideMargin(options.slideMargin);
  const charWeight = normalizeTableAutoPageWeight(
    options.autoPageCharWeight,
    'Table autoPageCharWeight',
  );
  const lineWeight = normalizeTableAutoPageWeight(
    options.autoPageLineWeight,
    'Table autoPageLineWeight',
  );
  const measureContent = context.autoRowHeight
    || context.rowHeights.some((height) => height === 0)
    || charWeight !== undefined
    || lineWeight !== undefined
    || context.hasCellMeasurementWeights;

  return Object.freeze({
    repeatHeader,
    headerRows,
    ...(slideStartY === undefined ? {} : { slideStartY }),
    ...(slideMargin === undefined ? {} : { slideMargin }),
    ...(charWeight === undefined ? {} : { charWeight }),
    ...(lineWeight === undefined ? {} : { lineWeight }),
    measureContent,
  });
}

export function normalizeTableAutoPageEnabled(value: unknown): boolean {
  if (value !== undefined && typeof value !== 'boolean') {
    throw new TypeError('Table autoPage must be a boolean');
  }
  return value === true;
}

export function normalizeTableAutoPageWeight(
  value: unknown,
  context: string,
): number | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  if (value < -1 || value > 1) {
    throw new RangeError(`${context} must be between -1 and 1`);
  }
  return Object.is(value, -0) ? 0 : value;
}

export function resolveTableAutoPageLayout(
  definition: Readonly<NormalizedTableDefinition>,
  slideSize: Readonly<SlideSize>,
  layoutMargins?: Readonly<TableAutoPageMargins>,
  bottomEdgeOverride?: number,
): Readonly<TableAutoPageLayoutRegion> {
  if (definition.autoPage === undefined) {
    throw new TypeError('Table auto-page layout requires an auto-page definition');
  }
  const width = normalizePositiveSafeInteger(slideSize.width, 'Slide width');
  const height = normalizePositiveSafeInteger(slideSize.height, 'Slide height');
  const margins = definition.autoPage.slideMargin === undefined
    ? layoutMargins === undefined
      ? canonicalMargins()
      : normalizeLayoutMargins(layoutMargins)
    : tupleMargins(definition.autoPage.slideMargin);
  const horizontalMargins = safeSum(
    [margins.left, margins.right],
    'Table auto-page horizontal margins',
  );
  const verticalMargins = safeSum(
    [margins.top, margins.bottom],
    'Table auto-page vertical margins',
  );
  if (horizontalMargins >= width) {
    throw new RangeError('Table auto-page horizontal margins must be smaller than slide width');
  }
  if (verticalMargins >= height) {
    throw new RangeError('Table auto-page vertical margins must be smaller than slide height');
  }

  const canonicalBottomEdge = height - margins.bottom;
  const bottomEdge = bottomEdgeOverride === undefined
    ? canonicalBottomEdge
    : Math.min(
        canonicalBottomEdge,
        normalizePositiveSafeInteger(
          bottomEdgeOverride,
          'Table auto-page bottom edge override',
        ),
      );
  const firstY = normalizeNonNegativeSafeInteger(
    definition.y,
    'Table auto-page source Y',
  );
  const firstCapacity = bottomEdge - firstY;
  if (!Number.isSafeInteger(firstCapacity) || firstCapacity <= 0) {
    throw new RangeError('Table auto-page source height must be positive');
  }
  const continuationY = definition.autoPage.slideStartY ?? margins.top;
  if (!Number.isSafeInteger(continuationY) || continuationY < 0) {
    throw new RangeError('Table auto-page continuation Y must be a non-negative safe integer');
  }
  const continuationCapacity = bottomEdge - continuationY;
  if (!Number.isSafeInteger(continuationCapacity) || continuationCapacity <= 0) {
    throw new RangeError('Table auto-page continuation height must be positive');
  }

  return Object.freeze({
    firstY,
    continuationY,
    bottomEdge,
    firstCapacity,
    continuationCapacity,
  });
}

export function planTableAutoPages(
  definition: Readonly<NormalizedTableDefinition>,
  region: Readonly<TableAutoPageLayoutRegion>,
): readonly Readonly<NormalizedTableDefinition>[] {
  if (definition.autoPage === undefined) return Object.freeze([definition]);
  if (definition.autoPage.measureContent) {
    throw new RangeError('Table auto-page rowHeights must be materialized before planning');
  }
  const {
    firstY,
    continuationY,
    firstCapacity,
    continuationCapacity,
  } = region;

  const headerRows = definition.autoPage.headerRows;
  assertHeaderBoundary(definition.rows, headerRows);
  const headerIndexes = Array.from({ length: headerRows }, (_, index) => index);
  const headerHeight = sumSelectedHeights(
    definition.rowHeights,
    headerIndexes,
    'Table auto-page header height',
  );
  if (headerHeight > firstCapacity) {
    throw new RangeError('Table auto-page headers do not fit on the source slide');
  }
  const blocks = tableRowBlocks(definition, headerRows);
  for (const block of blocks) {
    if (headerHeight + block.height > continuationCapacity) {
      throw new RangeError('Table auto-page row block does not fit on a continuation slide');
    }
  }

  const pages: number[][] = [];
  let current = [...headerIndexes];
  let currentHeight = headerHeight;
  let capacity = firstCapacity;
  for (const block of blocks) {
    if (currentHeight + block.height <= capacity) {
      appendRange(current, block.start, block.end);
      currentHeight += block.height;
      continue;
    }
    if (current.length === 0) {
      throw new RangeError('Table auto-page first row block does not fit on the source slide');
    }
    pages.push(current);
    current = [...headerIndexes];
    currentHeight = headerHeight;
    capacity = continuationCapacity;
    appendRange(current, block.start, block.end);
    currentHeight += block.height;
  }
  if (current.length === 0) {
    throw new RangeError('Table auto-page source slide must contain at least one row');
  }
  pages.push(current);

  return Object.freeze(pages.map((rowIndexes, pageIndex) => pageDefinition(
    definition,
    rowIndexes,
    pageIndex === 0 ? firstY : continuationY,
  )));
}

function normalizeHeaderRows(value: unknown, rowCount: number): number {
  if (value === undefined) return 1;
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError('Table autoPageHeaderRows must be finite');
  }
  if (!Number.isSafeInteger(value) || value < 1 || value > rowCount) {
    throw new RangeError(
      'Table autoPageHeaderRows must be a positive safe integer within the table row count',
    );
  }
  return value;
}

function normalizeSlideMargin(
  value: unknown,
): readonly [number, number, number, number] {
  const values = typeof value === 'number'
    ? [value, value, value, value]
    : readMarginTuple(value);
  return Object.freeze(values.map((side, index) =>
    normalizeNonNegativeSafeInteger(
      side,
      `Table slideMargin ${['top', 'right', 'bottom', 'left'][index]}`,
    )) as [number, number, number, number]);
}

function readMarginTuple(value: unknown): readonly unknown[] {
  if (!Array.isArray(value)) {
    throw new TypeError('Table slideMargin must be a number or four-item array');
  }
  if (Object.getPrototypeOf(value) !== Array.prototype) {
    throw new TypeError('Table slideMargin must be a plain array');
  }
  if (value.length !== 4) {
    throw new RangeError('Table slideMargin must contain exactly four values');
  }
  const allowed = new Set(['length', '0', '1', '2', '3']);
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !allowed.has(key)) {
      throw new TypeError(`Table slideMargin contains unsupported property ${String(key)}`);
    }
  }
  return Array.from({ length: 4 }, (_, index) => {
    const descriptor = Object.getOwnPropertyDescriptor(value, String(index));
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError('Table slideMargin must be dense and contain only data items');
    }
    return descriptor.value;
  });
}

function normalizeNonNegativeSafeInteger(value: unknown, context: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  if (!Number.isSafeInteger(value)) {
    throw new TypeError(`${context} must be a safe integer EMU value`);
  }
  if (value < 0) throw new RangeError(`${context} must be non-negative`);
  return value;
}

function normalizePositiveSafeInteger(value: unknown, context: string): number {
  const normalized = normalizeNonNegativeSafeInteger(value, context);
  if (normalized <= 0) throw new RangeError(`${context} must be positive`);
  return normalized;
}

function canonicalMargins(): Readonly<TableAutoPageMargins> {
  return Object.freeze({
    top: DEFAULT_PAGE_MARGIN,
    right: DEFAULT_PAGE_MARGIN,
    bottom: DEFAULT_PAGE_MARGIN,
    left: DEFAULT_PAGE_MARGIN,
  });
}

function tupleMargins(
  value: readonly [number, number, number, number],
): Readonly<TableAutoPageMargins> {
  return Object.freeze({
    top: value[0],
    right: value[1],
    bottom: value[2],
    left: value[3],
  });
}

function normalizeLayoutMargins(
  value: Readonly<TableAutoPageMargins>,
): Readonly<TableAutoPageMargins> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError('Table auto-page layout margins must be an object');
  }
  return Object.freeze({
    top: normalizeNonNegativeSafeInteger(value.top, 'Table auto-page layout top margin'),
    right: normalizeNonNegativeSafeInteger(value.right, 'Table auto-page layout right margin'),
    bottom: normalizeNonNegativeSafeInteger(value.bottom, 'Table auto-page layout bottom margin'),
    left: normalizeNonNegativeSafeInteger(value.left, 'Table auto-page layout left margin'),
  });
}

function assertHeaderBoundary(
  rows: readonly (readonly NormalizedTableCell[])[],
  headerRows: number,
): void {
  for (let rowIndex = 0; rowIndex < headerRows; rowIndex += 1) {
    for (const cell of rows[rowIndex]!) {
      if (cell.rowspan !== undefined && rowIndex + cell.rowspan > headerRows) {
        throw new RangeError('Table auto-page merge cannot cross the repeated-header boundary');
      }
    }
  }
}

function tableRowBlocks(
  definition: Readonly<NormalizedTableDefinition>,
  start: number,
): readonly TableRowBlock[] {
  const blocks: TableRowBlock[] = [];
  let blockStart = start;
  while (blockStart < definition.rows.length) {
    let blockEnd = blockStart + 1;
    for (let rowIndex = blockStart; rowIndex < blockEnd; rowIndex += 1) {
      for (const cell of definition.rows[rowIndex]!) {
        if (cell.rowspan !== undefined) {
          blockEnd = Math.max(blockEnd, rowIndex + cell.rowspan);
        }
      }
    }
    if (blockEnd > definition.rows.length) {
      throw new RangeError('Table auto-page merge row block is out of range');
    }
    const indexes = Array.from(
      { length: blockEnd - blockStart },
      (_, index) => blockStart + index,
    );
    blocks.push(Object.freeze({
      start: blockStart,
      end: blockEnd,
      height: sumSelectedHeights(
        definition.rowHeights,
        indexes,
        'Table auto-page row block height',
      ),
    }));
    blockStart = blockEnd;
  }
  return Object.freeze(blocks);
}

function appendRange(target: number[], start: number, end: number): void {
  for (let index = start; index < end; index += 1) target.push(index);
}

function pageDefinition(
  source: Readonly<NormalizedTableDefinition>,
  rowIndexes: readonly number[],
  y: number,
): Readonly<NormalizedTableDefinition> {
  const rows = Object.freeze(rowIndexes.map((index) => source.rows[index]!));
  const rowHeights = Object.freeze(rowIndexes.map((index) => source.rowHeights[index]!));
  const height = safeSum(rowHeights, 'Table auto-page page height');
  return Object.freeze({
    rows,
    ...(source.name === undefined ? {} : { name: source.name }),
    ...(source.placeholder === undefined ? {} : { placeholder: source.placeholder }),
    x: source.x,
    y,
    width: source.width,
    height,
    autoRowHeight: false,
    columnWidths: source.columnWidths,
    rowHeights,
  });
}

function sumSelectedHeights(
  heights: readonly number[],
  indexes: readonly number[],
  context: string,
): number {
  return safeSum(indexes.map((index) => {
    const height = heights[index];
    if (height === undefined) throw new RangeError(`${context} row is out of range`);
    return normalizePositiveSafeInteger(height, `${context} row ${index}`);
  }), context);
}

function safeSum(values: readonly number[], context: string): number {
  return values.reduce((sum, value) => {
    if (value > Number.MAX_SAFE_INTEGER - sum) {
      throw new RangeError(`${context} must fit a safe integer EMU value`);
    }
    return sum + value;
  }, 0);
}
