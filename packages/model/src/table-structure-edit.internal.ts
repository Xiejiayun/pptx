import type {
  LosslessXmlDocument,
  XmlAttribute,
  XmlElement,
} from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
import { renderEmptyTableCellFragment } from './table-create.internal.js';
import {
  readTableMergeStateFromCells,
  replaceTableCellMergeAttributes,
  type TableCellMergeTokens,
  type TableMergeRegionState,
  type TableMergeState,
} from './table-cell-merge.internal.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const MAX_TABLE_PHYSICAL_CELLS = 1_000_000;
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

export function insertTableRows(
  xml: LosslessXmlDocument,
  frame: XmlElement,
  input: Readonly<NormalizedTableRowInsert>,
  partUri: string,
): void {
  const structure = requireEditableTableStructure(frame, partUri);
  const rowCount = structure.rows.length;
  const columnCount = structure.gridColumns.length;
  if (input.rowIndex > rowCount) {
    throw new RangeError(`Table row insert index ${input.rowIndex} is out of range`);
  }
  const currentCells = rowCount * columnCount;
  const availableCells = MAX_TABLE_PHYSICAL_CELLS - currentCells;
  if (
    availableCells < 0
    || input.count > Math.floor(availableCells / columnCount)
  ) {
    throw new RangeError('Table row insert exceeds the physical table cell limit');
  }

  const insertedHeights = materializeInsertedDimensions(
    input.rowHeights,
    input.count,
    structure.rowHeights[input.rowIndex] ?? structure.rowHeights.at(-1)!,
    'Table row insert heights',
  );
  const finalHeights = [
    ...structure.rowHeights.slice(0, input.rowIndex),
    ...insertedHeights,
    ...structure.rowHeights.slice(input.rowIndex),
  ];
  const finalHeight = explicitDimensionTotal(finalHeights, 'Table row heights');
  const regionUpdates = structure.mergeState.regions.map((region) => {
    if (input.rowIndex <= region.rowIndex) {
      return {
        source: region,
        target: Object.freeze({
          ...region,
          rowIndex: region.rowIndex + input.count,
        }),
        changed: false,
      };
    }
    if (input.rowIndex < region.rowIndex + region.rowspan) {
      return {
        source: region,
        target: Object.freeze({
          ...region,
          rowspan: region.rowspan + input.count,
        }),
        changed: true,
      };
    }
    return { source: region, target: region, changed: false };
  });
  const finalRegions = regionUpdates.map(({ target }) => target);
  const mergeOwnership = validateMergeRegions(
    finalRegions,
    rowCount + input.count,
    columnCount,
  );

  const drawingDeclaration = namespaceUriForPrefix(structure.table, 'a')
    === DRAWING_NAMESPACE
    ? ''
    : ` xmlns:a="${DRAWING_NAMESPACE}"`;
  const insertedRows = insertedHeights.map((height, offset) => {
    const rowIndex = input.rowIndex + offset;
    const cells = Array.from({ length: columnCount }, (_, columnIndex) =>
      renderEmptyTableCellFragment(
        mergeTokensAt(mergeOwnership, columnCount, rowIndex, columnIndex),
      )).join('');
    return `<a:tr${drawingDeclaration} h="${height}">${cells}</a:tr>`;
  }).join('');

  for (const update of regionUpdates) {
    if (!update.changed) continue;
    for (let rowOffset = 0; rowOffset < update.source.rowspan; rowOffset += 1) {
      for (
        let columnOffset = 0;
        columnOffset < update.source.colspan;
        columnOffset += 1
      ) {
        const sourceRow = update.source.rowIndex + rowOffset;
        const sourceColumn = update.source.columnIndex + columnOffset;
        const targetRow = sourceRow < input.rowIndex
          ? sourceRow
          : sourceRow + input.count;
        replaceTableCellMergeAttributes(
          xml,
          structure.cells[sourceRow]![sourceColumn]!,
          mergeTokensForRegionCell(update.target, targetRow, sourceColumn),
        );
      }
    }
  }

  const insertionPoint = input.rowIndex < rowCount
    ? structure.rows[input.rowIndex]!.start
    : structure.rows.at(-1)!.end;
  xml.replace(insertionPoint, insertionPoint, insertedRows);
  if (finalHeight !== undefined && finalHeight !== structure.height) {
    xml.replaceAttribute(structure.heightAttribute, String(finalHeight));
  }
}

export function deleteTableRows(
  xml: LosslessXmlDocument,
  frame: XmlElement,
  input: Readonly<NormalizedTableDelete>,
  partUri: string,
): ReadonlySet<string> {
  const structure = requireEditableTableStructure(frame, partUri);
  const rowCount = structure.rows.length;
  const columnCount = structure.gridColumns.length;
  if (
    input.index >= rowCount
    || input.count > rowCount - input.index
  ) {
    throw new RangeError(
      `Table row delete range ${input.index}:${input.count} is out of range`,
    );
  }
  if (input.count === rowCount) {
    throw new RangeError('Table row delete must leave at least one row');
  }
  const deleteEnd = input.index + input.count;
  const finalHeights = structure.rowHeights.filter(
    (_, index) => index < input.index || index >= deleteEnd,
  );
  const finalHeight = explicitDimensionTotal(finalHeights, 'Table row heights');
  const regionUpdates = structure.mergeState.regions.map((region) => {
    const regionEnd = region.rowIndex + region.rowspan;
    const overlap = Math.max(
      0,
      Math.min(regionEnd, deleteEnd) - Math.max(region.rowIndex, input.index),
    );
    if (overlap === 0) {
      const target = region.rowIndex >= deleteEnd
        ? Object.freeze({ ...region, rowIndex: region.rowIndex - input.count })
        : region;
      return { source: region, target, changed: false, survivorCount: region.rowspan };
    }
    const survivorCount = region.rowspan - overlap;
    if (survivorCount === 0) {
      return { source: region, target: undefined, changed: true, survivorCount };
    }
    const target = Object.freeze({
      ...region,
      rowIndex: region.rowIndex < input.index ? region.rowIndex : input.index,
      rowspan: survivorCount,
    });
    return { source: region, target, changed: true, survivorCount };
  });
  const finalRegions = regionUpdates.flatMap(({ target }) =>
    target && (target.rowspan > 1 || target.colspan > 1) ? [target] : []);
  validateMergeRegions(finalRegions, rowCount - input.count, columnCount);

  const removedRelationshipIds = new Set<string>();
  for (let rowIndex = input.index; rowIndex < deleteEnd; rowIndex += 1) {
    collectRelationshipIds(structure.rows[rowIndex]!, removedRelationshipIds);
  }

  for (const update of regionUpdates) {
    if (!update.changed || update.survivorCount === 0) continue;
    for (let rowOffset = 0; rowOffset < update.source.rowspan; rowOffset += 1) {
      const sourceRow = update.source.rowIndex + rowOffset;
      if (sourceRow >= input.index && sourceRow < deleteEnd) continue;
      const targetRow = sourceRow < input.index
        ? sourceRow
        : sourceRow - input.count;
      for (
        let columnOffset = 0;
        columnOffset < update.source.colspan;
        columnOffset += 1
      ) {
        const columnIndex = update.source.columnIndex + columnOffset;
        const tokens = update.target
          && (update.target.rowspan > 1 || update.target.colspan > 1)
          ? mergeTokensForRegionCell(update.target, targetRow, columnIndex)
          : {};
        replaceTableCellMergeAttributes(
          xml,
          structure.cells[sourceRow]![columnIndex]!,
          tokens,
        );
      }
    }
  }

  for (let rowIndex = input.index; rowIndex < deleteEnd; rowIndex += 1) {
    const row = structure.rows[rowIndex]!;
    xml.replace(row.start, row.end, '');
  }
  if (finalHeight !== undefined && finalHeight !== structure.height) {
    xml.replaceAttribute(structure.heightAttribute, String(finalHeight));
  }
  return removedRelationshipIds;
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

function materializeInsertedDimensions(
  explicit: readonly number[] | undefined,
  count: number,
  fallback: number,
  context: string,
): readonly number[] {
  if (explicit === undefined) {
    return Array.from({ length: count }, () => fallback);
  }
  if (explicit.length !== count) {
    throw new TypeError(`${context} must match the insert count`);
  }
  return explicit;
}

function explicitDimensionTotal(
  values: readonly number[],
  context: string,
): number | undefined {
  if (values.some((value) => value === 0)) return undefined;
  let total = 0;
  for (const value of values) {
    if (value > Number.MAX_SAFE_INTEGER - total) {
      throw new RangeError(`${context} sum must fit a safe integer EMU value`);
    }
    total += value;
  }
  return total;
}

function mergeTokensAt(
  ownership: ReadonlyMap<number, Readonly<TableMergeRegionState>>,
  columnCount: number,
  rowIndex: number,
  columnIndex: number,
): Readonly<TableCellMergeTokens> {
  const region = ownership.get(rowIndex * columnCount + columnIndex);
  return region ? mergeTokensForRegionCell(region, rowIndex, columnIndex) : {};
}

function mergeTokensForRegionCell(
  region: Readonly<TableMergeRegionState>,
  rowIndex: number,
  columnIndex: number,
): Readonly<TableCellMergeTokens> {
  const rowOffset = rowIndex - region.rowIndex;
  const columnOffset = columnIndex - region.columnIndex;
  return {
    ...(rowOffset === 0 && region.rowspan > 1
      ? { rowSpan: region.rowspan }
      : {}),
    ...(columnOffset === 0 && region.colspan > 1
      ? { gridSpan: region.colspan }
      : {}),
    ...(rowOffset > 0 ? { vertical: true as const } : {}),
    ...(columnOffset > 0 ? { horizontal: true as const } : {}),
  };
}

function validateMergeRegions(
  regions: readonly Readonly<TableMergeRegionState>[],
  rowCount: number,
  columnCount: number,
): ReadonlyMap<number, Readonly<TableMergeRegionState>> {
  const ownership = new Map<number, Readonly<TableMergeRegionState>>();
  for (const region of regions) {
    if (
      region.rowspan < 1
      || region.colspan < 1
      || (region.rowspan === 1 && region.colspan === 1)
      || region.rowIndex < 0
      || region.columnIndex < 0
      || region.rowspan > rowCount - region.rowIndex
      || region.colspan > columnCount - region.columnIndex
    ) {
      throw new RangeError('Projected table merge region is out of range');
    }
    for (let rowOffset = 0; rowOffset < region.rowspan; rowOffset += 1) {
      for (let columnOffset = 0; columnOffset < region.colspan; columnOffset += 1) {
        const key = (region.rowIndex + rowOffset) * columnCount
          + region.columnIndex + columnOffset;
        if (ownership.has(key)) {
          throw new RangeError('Projected table merge regions overlap');
        }
        ownership.set(key, region);
      }
    }
  }
  return ownership;
}

function collectRelationshipIds(
  root: XmlElement,
  relationshipIds: Set<string>,
): void {
  for (const attribute of root.attributes) {
    const prefix = lexicalPrefix(attribute.name);
    if (
      prefix !== ''
      && attribute.value.length > 0
      && namespaceUriForPrefix(root, prefix) === RELATIONSHIP_NAMESPACE
    ) relationshipIds.add(attribute.value);
  }
  for (const child of root.children) {
    if (child.type === 'element') collectRelationshipIds(child, relationshipIds);
  }
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
  return namespaceUriForPrefix(element, lexicalPrefix(element.name));
}

function namespaceUriForPrefix(
  element: XmlElement,
  prefix: string,
): string | undefined {
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
