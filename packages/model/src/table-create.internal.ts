import { escapeXmlAttribute } from '@pptx/lossless-xml';
import {
  normalizeRichText,
  renderRichTextParagraphs,
} from './rich-text.internal.js';

const EMU_PER_INCH = 914_400;
const DEFAULT_OFFSET = EMU_PER_INCH / 2;
const DEFAULT_HEIGHT = EMU_PER_INCH;
const CELL_MARGIN_HORIZONTAL = 91_440;
const CELL_MARGIN_VERTICAL = 45_720;
const OPTION_KEYS = [
  'name',
  'x',
  'y',
  'width',
  'height',
  'columnWidths',
  'rowHeights',
] as const;
const NO_BORDERS = ['lnL', 'lnR', 'lnT', 'lnB']
  .map((tag) => `<a:${tag} w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:${tag}>`)
  .join('');

export interface NormalizedTableDefinition {
  readonly rows: readonly (readonly string[])[];
  readonly name?: string;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly autoRowHeight: boolean;
  readonly columnWidths: readonly number[];
  readonly rowHeights: readonly number[];
}

export function normalizeTableDefinition(
  rows: unknown,
  options: unknown,
): NormalizedTableDefinition {
  const outer = readDenseArray(rows, 'Table rows');
  const normalizedRows = outer.map((row, rowIndex) =>
    readDenseArray(row, `Table row ${rowIndex}`).map((cell, columnIndex) => {
      if (typeof cell !== 'string') {
        throw new TypeError(`Table cell ${rowIndex},${columnIndex} must be a string`);
      }
      if (/\r|\n/.test(cell)) {
        throw new TypeError(`Table cell ${rowIndex},${columnIndex} must contain one paragraph`);
      }
      if (containsInvalidXmlCharacter(cell)) {
        throw new TypeError(`Table cell ${rowIndex},${columnIndex} contains invalid XML characters`);
      }
      return cell;
    }));
  const columnCount = normalizedRows[0]!.length;
  for (let rowIndex = 1; rowIndex < normalizedRows.length; rowIndex += 1) {
    if (normalizedRows[rowIndex]!.length !== columnCount) {
      throw new TypeError('Table rows must contain the same number of cells');
    }
  }

  const normalizedOptions = readOptions(options);
  const name = normalizedOptions.name;
  if (name !== undefined) {
    if (typeof name !== 'string') throw new TypeError('Table name must be a string');
    if (containsInvalidXmlCharacter(name)) {
      throw new TypeError('Table name contains invalid XML characters');
    }
  }

  const x = normalizedOptions.x === undefined
    ? DEFAULT_OFFSET
    : normalizeCoordinate(normalizedOptions.x, 'Table x');
  const y = normalizedOptions.y === undefined
    ? DEFAULT_OFFSET
    : normalizeCoordinate(normalizedOptions.y, 'Table y');
  const requestedColumnWidths = normalizedOptions.columnWidths;
  let width: number;
  let columnWidths: readonly number[];
  if (requestedColumnWidths !== undefined) {
    columnWidths = normalizeDimensionVector(
      requestedColumnWidths,
      columnCount,
      'Table columnWidths',
      'column count',
    );
    const columnWidthSum = sumDimensions(columnWidths, 'Table columnWidths');
    if (normalizedOptions.width === undefined) {
      width = columnWidthSum;
    } else {
      width = normalizeCoordinate(normalizedOptions.width, 'Table width');
      if (width !== columnWidthSum) {
        throw new RangeError('Table width must equal the sum of columnWidths');
      }
    }
  } else {
    const defaultWidth = columnCount * EMU_PER_INCH;
    if (!Number.isSafeInteger(defaultWidth)) {
      throw new RangeError('Table default width must fit a safe integer EMU value');
    }
    width = normalizedOptions.width === undefined
      ? defaultWidth
      : normalizeCoordinate(normalizedOptions.width, 'Table width');
    if (width < columnCount) {
      throw new RangeError('Table width must provide at least one EMU per column');
    }
    columnWidths = distributeTableDimension(width, columnCount);
  }

  const requestedRowHeights = normalizedOptions.rowHeights;
  let height: number;
  let autoRowHeight: boolean;
  let rowHeights: readonly number[];
  if (requestedRowHeights !== undefined) {
    rowHeights = normalizeDimensionVector(
      requestedRowHeights,
      normalizedRows.length,
      'Table rowHeights',
      'row count',
    );
    const rowHeightSum = sumDimensions(rowHeights, 'Table rowHeights');
    autoRowHeight = false;
    if (normalizedOptions.height === undefined) {
      height = rowHeightSum;
    } else {
      height = normalizeCoordinate(normalizedOptions.height, 'Table height');
      if (height !== rowHeightSum) {
        throw new RangeError('Table height must equal the sum of rowHeights');
      }
    }
  } else {
    autoRowHeight = normalizedOptions.height === undefined;
    height = autoRowHeight
      ? DEFAULT_HEIGHT
      : normalizeCoordinate(normalizedOptions.height, 'Table height');
    if (height <= 0) throw new RangeError('Table height must be greater than zero');
    if (!autoRowHeight && height < normalizedRows.length) {
      throw new RangeError('Table height must provide at least one EMU per row');
    }
    rowHeights = autoRowHeight
      ? normalizedRows.map(() => 0)
      : distributeTableDimension(height, normalizedRows.length);
  }

  return {
    rows: normalizedRows,
    ...(name !== undefined ? { name } : {}),
    x,
    y,
    width,
    height,
    autoRowHeight,
    columnWidths,
    rowHeights,
  };
}

export function distributeTableDimension(total: number, count: number): readonly number[] {
  const quotient = Math.floor(total / count);
  const remainder = total % count;
  return Array.from(
    { length: count },
    (_, index) => quotient + (index < remainder ? 1 : 0),
  );
}

export function renderTableGraphicFrame(
  id: number,
  definition: NormalizedTableDefinition,
): string {
  const grid = definition.columnWidths.map((width) => `<a:gridCol w="${width}"/>`).join('');
  const rows = definition.rows.map((row, rowIndex) => {
    const cells = row.map(renderTableCell).join('');
    return `<a:tr h="${definition.rowHeights[rowIndex]}">${cells}</a:tr>`;
  }).join('');
  const name = escapeXmlAttribute(definition.name ?? `Table ${id}`);

  return `<p:graphicFrame xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:nvGraphicFramePr><p:cNvPr id="${id}" name="${name}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${definition.x}" y="${definition.y}"/><a:ext cx="${definition.width}" cy="${definition.height}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/><a:tblGrid>${grid}</a:tblGrid>${rows}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
}

function readDenseArray(value: unknown, context: string): readonly unknown[] {
  if (!Array.isArray(value) || value.length === 0) {
    throw new TypeError(`${context} must be a non-empty array`);
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

function readOptions(value: unknown): Record<string, unknown> {
  if (value === undefined) return {};
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError('Table options must be an object');
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError('Table options must be an ordinary object');
  }
  const allowed = new Set<string>(OPTION_KEYS);
  const result: Record<string, unknown> = {};
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !allowed.has(key)) {
      throw new TypeError(`Table options contain unsupported property ${String(key)}`);
    }
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`Table option ${key} must be a data property`);
    }
    result[key] = descriptor.value;
  }
  return result;
}

function normalizeCoordinate(value: unknown, context: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  const rounded = Math.round(value);
  if (!Number.isSafeInteger(rounded)) {
    throw new RangeError(`${context} must round to a safe integer EMU value`);
  }
  return rounded;
}

function normalizeDimensionVector(
  value: unknown,
  count: number,
  context: string,
  countContext: string,
): readonly number[] {
  const values = Array.isArray(value)
    ? readDenseArray(value, context)
    : Array.from({ length: count }, () => value);
  if (values.length !== count) {
    throw new TypeError(`${context} must match the table ${countContext}`);
  }
  return values.map((item, index) => {
    const dimension = normalizeCoordinate(item, `${context} ${index}`);
    if (dimension <= 0) {
      throw new RangeError(`${context} ${index} must be greater than zero`);
    }
    return dimension;
  });
}

function sumDimensions(dimensions: readonly number[], context: string): number {
  return dimensions.reduce((sum, dimension) => {
    if (dimension > Number.MAX_SAFE_INTEGER - sum) {
      throw new RangeError(`${context} sum must fit a safe integer EMU value`);
    }
    return sum + dimension;
  }, 0);
}

function renderTableCell(text: string): string {
  const paragraphs = renderRichTextParagraphs(normalizeRichText([
    { runs: [{ text, style: {} }] },
  ]));
  return `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>${paragraphs}</a:txBody><a:tcPr marL="${CELL_MARGIN_HORIZONTAL}" marR="${CELL_MARGIN_HORIZONTAL}" marT="${CELL_MARGIN_VERTICAL}" marB="${CELL_MARGIN_VERTICAL}">${NO_BORDERS}</a:tcPr></a:tc>`;
}

function containsInvalidXmlCharacter(value: string): boolean {
  return /[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(value);
}
