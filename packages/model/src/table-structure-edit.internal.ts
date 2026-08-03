import type {
  XmlAttribute,
  XmlElement,
} from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
import {
  readTableMergeStateFromCells,
  type TableMergeState,
} from './table-cell-merge.internal.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const MAX_INSERT_ITEMS = 999_999;
const ROW_INSERT_KEYS = ['count', 'rowHeights'] as const;
const COLUMN_INSERT_KEYS = ['count', 'columnWidths'] as const;

export interface NormalizedTableRowInsert {
  readonly rowIndex: number;
  readonly count: number;
  readonly rowHeights?: readonly number[];
}

export interface NormalizedTableColumnInsert {
  readonly columnIndex: number;
  readonly count: number;
  readonly columnWidths?: readonly number[];
}

export interface NormalizedTableDelete {
  readonly index: number;
  readonly count: number;
}

export interface EditableTableStructure {
  readonly table: XmlElement;
  readonly grid: XmlElement;
  readonly gridColumns: readonly XmlElement[];
  readonly rows: readonly XmlElement[];
  readonly cells: readonly (readonly XmlElement[])[];
  readonly columnWidths: readonly number[];
  readonly rowHeights: readonly number[];
  readonly widthAttribute: XmlAttribute;
  readonly heightAttribute: XmlAttribute;
  readonly width: number;
  readonly height: number;
  readonly mergeState: Readonly<TableMergeState>;
}

export function normalizeTableRowInsertInput(
  rowIndex: unknown,
  options: unknown,
): Readonly<NormalizedTableRowInsert> {
  const normalizedRowIndex = normalizeIndex(rowIndex, 'Table row insert index');
  const input = readOptions(options, 'Table row insert options', ROW_INSERT_KEYS);
  const count = normalizeInsertCount(input.count, 'Table row insert count');
  const rowHeights = input.rowHeights === undefined
    ? undefined
    : normalizeDimensions(
        input.rowHeights,
        count,
        'Table row insert heights',
        false,
      );
  return Object.freeze({
    rowIndex: normalizedRowIndex,
    count,
    ...(rowHeights === undefined ? {} : { rowHeights }),
  });
}

export function normalizeTableColumnInsertInput(
  columnIndex: unknown,
  options: unknown,
): Readonly<NormalizedTableColumnInsert> {
  const normalizedColumnIndex = normalizeIndex(
    columnIndex,
    'Table column insert index',
  );
  const input = readOptions(
    options,
    'Table column insert options',
    COLUMN_INSERT_KEYS,
  );
  const count = normalizeInsertCount(input.count, 'Table column insert count');
  const columnWidths = input.columnWidths === undefined
    ? undefined
    : normalizeDimensions(
        input.columnWidths,
        count,
        'Table column insert widths',
        true,
      );
  return Object.freeze({
    columnIndex: normalizedColumnIndex,
    count,
    ...(columnWidths === undefined ? {} : { columnWidths }),
  });
}

export function normalizeTableDeleteInput(
  index: unknown,
  count: unknown,
  axis: 'row' | 'column',
): Readonly<NormalizedTableDelete> {
  const context = axis === 'row' ? 'Table row delete' : 'Table column delete';
  return Object.freeze({
    index: normalizeIndex(index, `${context} index`),
    count: normalizeCount(count, `${context} count`),
  });
}

export function requireEditableTableStructure(
  frame: XmlElement,
  partUri: string,
): Readonly<EditableTableStructure> {
  const invalid = (): never => {
    throw new ModelParseError('Table structure is not safely editable', partUri);
  };
  if (
    frame.localName !== 'graphicFrame'
    || namespaceUri(frame) !== PRESENTATION_NAMESPACE
  ) return invalid();

  const transform = exactDirectChild(frame, 'xfrm', PRESENTATION_NAMESPACE);
  const extent = transform
    ? exactDirectChild(transform, 'ext', DRAWING_NAMESPACE)
    : undefined;
  const widthToken = extent ? readIntegerAttribute(extent, 'cx', false) : undefined;
  const heightToken = extent ? readIntegerAttribute(extent, 'cy', false) : undefined;
  if (!widthToken || !heightToken) return invalid();

  const graphic = exactDirectChild(frame, 'graphic', DRAWING_NAMESPACE);
  const graphicData = graphic
    ? exactDirectChild(graphic, 'graphicData', DRAWING_NAMESPACE)
    : undefined;
  const table = graphicData
    ? exactDirectChild(graphicData, 'tbl', DRAWING_NAMESPACE)
    : undefined;
  const grid = table
    ? exactDirectChild(table, 'tblGrid', DRAWING_NAMESPACE)
    : undefined;
  if (!table || !grid) return invalid();

  const gridColumns = directChildren(grid, 'gridCol', DRAWING_NAMESPACE);
  const rows = directChildren(table, 'tr', DRAWING_NAMESPACE);
  if (gridColumns.length === 0 || rows.length === 0) return invalid();

  const columnWidths: number[] = [];
  let widthSum = 0;
  for (const column of gridColumns) {
    const token = readIntegerAttribute(column, 'w', true);
    if (!token || token.value > Number.MAX_SAFE_INTEGER - widthSum) {
      return invalid();
    }
    widthSum += token.value;
    columnWidths.push(token.value);
  }

  const cells: XmlElement[][] = [];
  const rowHeights: number[] = [];
  for (const row of rows) {
    const token = readIntegerAttribute(row, 'h', false);
    const rowCells = directChildren(row, 'tc', DRAWING_NAMESPACE);
    if (!token || rowCells.length !== gridColumns.length) return invalid();
    rowHeights.push(token.value);
    cells.push(rowCells);
  }

  const mergeState = readTableMergeStateFromCells(cells);
  if (!mergeState) return invalid();
  return Object.freeze({
    table,
    grid,
    gridColumns: Object.freeze([...gridColumns]),
    rows: Object.freeze([...rows]),
    cells: Object.freeze(cells.map((row) => Object.freeze([...row]))),
    columnWidths: Object.freeze(columnWidths),
    rowHeights: Object.freeze(rowHeights),
    widthAttribute: widthToken.attribute,
    heightAttribute: heightToken.attribute,
    width: widthToken.value,
    height: heightToken.value,
    mergeState,
  });
}

function normalizeIndex(value: unknown, context: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  if (!Number.isSafeInteger(value) || value < 0) {
    throw new RangeError(`${context} must be a non-negative safe integer`);
  }
  return value;
}

function normalizeCount(value: unknown, context: string): number {
  if (value === undefined) return 1;
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  if (!Number.isSafeInteger(value) || value < 1) {
    throw new RangeError(`${context} must be a positive safe integer`);
  }
  return value;
}

function normalizeInsertCount(value: unknown, context: string): number {
  const count = normalizeCount(value, context);
  if (count > MAX_INSERT_ITEMS) {
    throw new RangeError(`${context} exceeds the physical table cell limit`);
  }
  return count;
}

function normalizeDimensions(
  value: unknown,
  count: number,
  context: string,
  positive: boolean,
): readonly number[] {
  const source = Array.isArray(value)
    ? readDenseArray(value, context)
    : Array.from({ length: count }, () => value);
  if (source.length !== count) {
    throw new TypeError(`${context} must match the insert count`);
  }
  return Object.freeze(source.map((item, index) =>
    normalizeDimension(item, `${context} ${index}`, positive)));
}

function normalizeDimension(
  value: unknown,
  context: string,
  positive: boolean,
): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  if (!positive && value < 0) {
    throw new RangeError(`${context} must be non-negative`);
  }
  const rounded = Math.round(value);
  if (!Number.isSafeInteger(rounded)) {
    throw new RangeError(`${context} must round to a safe integer EMU value`);
  }
  if (positive && rounded <= 0) {
    throw new RangeError(`${context} must be greater than zero`);
  }
  return rounded === 0 ? 0 : rounded;
}

function readOptions(
  value: unknown,
  context: string,
  supported: readonly string[],
): Record<string, unknown> {
  if (value === undefined) return Object.create(null) as Record<string, unknown>;
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${context} must be an object`);
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError(`${context} must be an ordinary or null-prototype object`);
  }
  const allowed = new Set(supported);
  const result: Record<string, unknown> = Object.create(null) as Record<string, unknown>;
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !allowed.has(key)) {
      throw new TypeError(`${context} contains unsupported property ${String(key)}`);
    }
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`${context} must contain only data properties`);
    }
    result[key] = descriptor.value;
  }
  return result;
}

function readDenseArray(
  value: readonly unknown[],
  context: string,
): readonly unknown[] {
  if (Object.getPrototypeOf(value) !== Array.prototype) {
    throw new TypeError(`${context} must be an ordinary array`);
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

function exactDirectChild(
  element: XmlElement,
  localName: string,
  namespace: string,
): XmlElement | undefined {
  const matches = directChildren(element, localName, namespace);
  return matches.length === 1 ? matches[0] : undefined;
}

function directChildren(
  element: XmlElement,
  localName: string,
  namespace: string,
): XmlElement[] {
  return element.children.filter((child): child is XmlElement =>
    child.type === 'element'
    && child.localName === localName
    && namespaceUri(child) === namespace);
}

function readIntegerAttribute(
  element: XmlElement,
  name: string,
  positive: boolean,
): { readonly attribute: XmlAttribute; readonly value: number } | undefined {
  const attributes = element.attributes.filter((attribute) => attribute.name === name);
  if (attributes.length !== 1) return undefined;
  const attribute = attributes[0]!;
  if (!/^[0-9]+$/u.test(attribute.value)) return undefined;
  const value = Number(attribute.value);
  if (!Number.isSafeInteger(value) || (positive ? value < 1 : value < 0)) {
    return undefined;
  }
  return { attribute, value };
}

function namespaceUri(element: XmlElement): string | undefined {
  const prefix = lexicalPrefix(element.name);
  const declaration = prefix === '' ? 'xmlns' : `xmlns:${prefix}`;
  for (let current: XmlElement | undefined = element; current; current = current.parent) {
    const attributes = current.attributes.filter(({ name }) => name === declaration);
    if (attributes.length > 1) return undefined;
    if (attributes[0]) return attributes[0].value;
  }
  return undefined;
}

function lexicalPrefix(name: string): string {
  const separator = name.indexOf(':');
  return separator < 0 ? '' : name.slice(0, separator);
}
