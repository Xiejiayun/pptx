import type {
  LosslessXmlDocument,
  XmlAttribute,
  XmlElement,
} from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';

const MERGE_ATTRIBUTE_NAMES = [
  'rowSpan',
  'gridSpan',
  'vMerge',
  'hMerge',
] as const;

export interface TableMergeRegionState {
  readonly rowIndex: number;
  readonly columnIndex: number;
  readonly rowspan: number;
  readonly colspan: number;
}

export interface TableCellMergeState extends TableMergeRegionState {
  readonly isAnchor: boolean;
}

export interface TableMergeState {
  readonly regions: readonly Readonly<TableMergeRegionState>[];
  readonly cells: readonly (readonly (Readonly<TableCellMergeState> | undefined)[])[];
}

interface DirectTableStructure {
  readonly columnCount: number;
  readonly cells: readonly (readonly XmlElement[])[];
}

interface ParsedMergeCell {
  readonly rowSpan: number;
  readonly gridSpan: number;
  readonly verticalContinuation: boolean;
  readonly horizontalContinuation: boolean;
}

interface LocalEdit {
  readonly start: number;
  readonly end: number;
  readonly replacement: string;
}

export function readTableMergeState(
  frame: XmlElement,
): Readonly<TableMergeState> | undefined {
  const structure = readDirectTableStructure(frame);
  if (!structure) return undefined;
  const parsed = structure.cells.map((row) => row.map(readMergeCell));
  if (parsed.some((row) => row.some((cell) => cell === undefined))) {
    return undefined;
  }
  const cells = parsed as readonly (readonly ParsedMergeCell[])[];
  const regions: TableMergeRegionState[] = [];
  for (let rowIndex = 0; rowIndex < cells.length; rowIndex += 1) {
    for (let columnIndex = 0; columnIndex < structure.columnCount; columnIndex += 1) {
      const cell = cells[rowIndex]![columnIndex]!;
      if (
        !cell.verticalContinuation
        && !cell.horizontalContinuation
        && (cell.rowSpan > 1 || cell.gridSpan > 1)
      ) {
        regions.push({
          rowIndex,
          columnIndex,
          rowspan: cell.rowSpan,
          colspan: cell.gridSpan,
        });
      }
    }
  }

  const ownership: (TableMergeRegionState | undefined)[][] = cells.map(() =>
    Array.from({ length: structure.columnCount }, () => undefined));
  for (const region of regions) {
    if (
      region.rowIndex + region.rowspan > cells.length
      || region.columnIndex + region.colspan > structure.columnCount
    ) return undefined;
    for (let rowOffset = 0; rowOffset < region.rowspan; rowOffset += 1) {
      for (let columnOffset = 0; columnOffset < region.colspan; columnOffset += 1) {
        const rowIndex = region.rowIndex + rowOffset;
        const columnIndex = region.columnIndex + columnOffset;
        if (ownership[rowIndex]![columnIndex] !== undefined) return undefined;
        ownership[rowIndex]![columnIndex] = region;
      }
    }
  }

  const snapshotCells: Array<
    readonly (Readonly<TableCellMergeState> | undefined)[]
  > = [];
  for (let rowIndex = 0; rowIndex < cells.length; rowIndex += 1) {
    const snapshotRow: (Readonly<TableCellMergeState> | undefined)[] = [];
    for (let columnIndex = 0; columnIndex < structure.columnCount; columnIndex += 1) {
      const cell = cells[rowIndex]![columnIndex]!;
      const region = ownership[rowIndex]![columnIndex];
      if (!region) {
        if (
          cell.rowSpan !== 1
          || cell.gridSpan !== 1
          || cell.verticalContinuation
          || cell.horizontalContinuation
        ) return undefined;
        snapshotRow.push(undefined);
        continue;
      }
      const rowOffset = rowIndex - region.rowIndex;
      const columnOffset = columnIndex - region.columnIndex;
      if (
        cell.rowSpan !== (rowOffset === 0 ? region.rowspan : 1)
        || cell.gridSpan !== (columnOffset === 0 ? region.colspan : 1)
        || cell.verticalContinuation !== (rowOffset > 0)
        || cell.horizontalContinuation !== (columnOffset > 0)
      ) return undefined;
      snapshotRow.push(Object.freeze({
        ...region,
        isAnchor: rowOffset === 0 && columnOffset === 0,
      }));
    }
    snapshotCells.push(Object.freeze(snapshotRow));
  }

  return Object.freeze({
    regions: Object.freeze(regions.map((region) => Object.freeze({ ...region }))),
    cells: Object.freeze(snapshotCells),
  });
}

export function normalizeTableMergeRegionInput(
  rowIndex: unknown,
  columnIndex: unknown,
  rowspan: unknown,
  colspan: unknown,
): Readonly<TableMergeRegionState> {
  const normalizedRowIndex = normalizeIndex(rowIndex, 'Table merge rowIndex');
  const normalizedColumnIndex = normalizeIndex(columnIndex, 'Table merge columnIndex');
  const normalizedRowspan = normalizeSpan(rowspan, 'Table merge rowspan');
  const normalizedColspan = normalizeSpan(colspan, 'Table merge colspan');
  if (normalizedRowspan === 1 && normalizedColspan === 1) {
    throw new RangeError('Table merge must cover more than one cell');
  }
  return Object.freeze({
    rowIndex: normalizedRowIndex,
    columnIndex: normalizedColumnIndex,
    rowspan: normalizedRowspan,
    colspan: normalizedColspan,
  });
}

export function replaceTableMergeRegion(
  xml: LosslessXmlDocument,
  frame: XmlElement,
  region: Readonly<TableMergeRegionState>,
  partUri: string,
): boolean {
  const state = requireEditableMergeState(frame, partUri);
  const rowCount = state.cells.length;
  const columnCount = state.cells[0]!.length;
  if (
    region.rowIndex >= rowCount
    || region.columnIndex >= columnCount
    || region.rowIndex + region.rowspan > rowCount
    || region.columnIndex + region.colspan > columnCount
  ) {
    throw new RangeError(
      `Table merge ${region.rowIndex},${region.columnIndex} ` +
      `${region.rowspan}x${region.colspan} is out of range`,
    );
  }
  const exact = state.regions.some((candidate) => regionsEqual(candidate, region));
  if (exact) return false;
  if (state.regions.some((candidate) => regionsOverlap(candidate, region))) {
    throw new RangeError('Table merge overlaps an existing merge region');
  }

  const structure = readDirectTableStructure(frame);
  if (!structure) {
    throw new ModelParseError('Table merge state is not safely editable', partUri);
  }
  for (let rowOffset = 0; rowOffset < region.rowspan; rowOffset += 1) {
    for (let columnOffset = 0; columnOffset < region.colspan; columnOffset += 1) {
      const cell = structure.cells[region.rowIndex + rowOffset]![
        region.columnIndex + columnOffset
      ]!;
      replaceCellMergeAttributes(xml, cell, {
        rowSpan: rowOffset === 0 ? region.rowspan : 1,
        gridSpan: columnOffset === 0 ? region.colspan : 1,
        verticalContinuation: rowOffset > 0,
        horizontalContinuation: columnOffset > 0,
      });
    }
  }
  return true;
}

export function clearTableMergeRegionAt(
  xml: LosslessXmlDocument,
  frame: XmlElement,
  rowIndex: unknown,
  columnIndex: unknown,
  partUri: string,
): boolean {
  const state = requireEditableMergeState(frame, partUri);
  const normalizedRowIndex = normalizeIndex(rowIndex, 'Table cell rowIndex');
  const normalizedColumnIndex = normalizeIndex(columnIndex, 'Table cell columnIndex');
  const row = state.cells[normalizedRowIndex];
  if (!row || normalizedColumnIndex >= row.length) {
    throw new RangeError(
      `Table cell ${normalizedRowIndex},${normalizedColumnIndex} was not found`,
    );
  }
  const region = row[normalizedColumnIndex];
  if (!region) return false;
  const structure = readDirectTableStructure(frame);
  if (!structure) {
    throw new ModelParseError('Table merge state is not safely editable', partUri);
  }
  for (let rowOffset = 0; rowOffset < region.rowspan; rowOffset += 1) {
    for (let columnOffset = 0; columnOffset < region.colspan; columnOffset += 1) {
      replaceCellMergeAttributes(
        xml,
        structure.cells[region.rowIndex + rowOffset]![region.columnIndex + columnOffset]!,
        {
          rowSpan: 1,
          gridSpan: 1,
          verticalContinuation: false,
          horizontalContinuation: false,
        },
      );
    }
  }
  return true;
}

function requireEditableMergeState(
  frame: XmlElement,
  partUri: string,
): Readonly<TableMergeState> {
  const state = readTableMergeState(frame);
  if (!state) {
    throw new ModelParseError('Table merge state is not safely editable', partUri);
  }
  return state;
}

function readDirectTableStructure(
  frame: XmlElement,
): DirectTableStructure | undefined {
  if (frame.localName !== 'graphicFrame') return undefined;
  const graphic = exactDirectChild(frame, 'graphic');
  const graphicData = graphic ? exactDirectChild(graphic, 'graphicData') : undefined;
  const table = graphicData ? exactDirectChild(graphicData, 'tbl') : undefined;
  if (!table) return undefined;
  const grids = directChildren(table, 'tblGrid');
  if (grids.length !== 1) return undefined;
  const columnCount = directChildren(grids[0]!, 'gridCol').length;
  if (columnCount === 0) return undefined;
  const rows = directChildren(table, 'tr');
  if (rows.length === 0) return undefined;
  const cells = rows.map((row) => directChildren(row, 'tc'));
  if (cells.some((row) => row.length !== columnCount)) return undefined;
  return { columnCount, cells };
}

function readMergeCell(cell: XmlElement): ParsedMergeCell | undefined {
  const rowSpan = readPositiveIntegerAttribute(cell, 'rowSpan');
  const gridSpan = readPositiveIntegerAttribute(cell, 'gridSpan');
  const verticalContinuation = readBooleanAttribute(cell, 'vMerge');
  const horizontalContinuation = readBooleanAttribute(cell, 'hMerge');
  if (
    rowSpan === undefined
    || gridSpan === undefined
    || verticalContinuation === undefined
    || horizontalContinuation === undefined
  ) return undefined;
  return {
    rowSpan,
    gridSpan,
    verticalContinuation,
    horizontalContinuation,
  };
}

function readPositiveIntegerAttribute(
  element: XmlElement,
  name: string,
): number | undefined {
  const attributes = exactAttributes(element, name);
  if (attributes.length > 1) return undefined;
  if (attributes.length === 0) return 1;
  const value = attributes[0]!.value;
  if (!/^[0-9]+$/u.test(value)) return undefined;
  const number = Number(value);
  return Number.isSafeInteger(number) && number > 0 ? number : undefined;
}

function readBooleanAttribute(
  element: XmlElement,
  name: string,
): boolean | undefined {
  const attributes = exactAttributes(element, name);
  if (attributes.length > 1) return undefined;
  if (attributes.length === 0) return false;
  const value = attributes[0]!.value;
  if (value === '1' || value === 'true') return true;
  if (value === '0' || value === 'false') return false;
  return undefined;
}

function exactAttributes(element: XmlElement, name: string): readonly XmlAttribute[] {
  return element.attributes.filter((attribute) => attribute.name === name);
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

function normalizeSpan(value: unknown, context: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  if (!Number.isSafeInteger(value) || value < 1) {
    throw new RangeError(`${context} must be a positive safe integer`);
  }
  return value;
}

function regionsEqual(
  left: Readonly<TableMergeRegionState>,
  right: Readonly<TableMergeRegionState>,
): boolean {
  return left.rowIndex === right.rowIndex
    && left.columnIndex === right.columnIndex
    && left.rowspan === right.rowspan
    && left.colspan === right.colspan;
}

function regionsOverlap(
  left: Readonly<TableMergeRegionState>,
  right: Readonly<TableMergeRegionState>,
): boolean {
  return left.rowIndex < right.rowIndex + right.rowspan
    && right.rowIndex < left.rowIndex + left.rowspan
    && left.columnIndex < right.columnIndex + right.colspan
    && right.columnIndex < left.columnIndex + left.colspan;
}

function replaceCellMergeAttributes(
  xml: LosslessXmlDocument,
  cell: XmlElement,
  state: ParsedMergeCell,
): void {
  const source = xml.source.slice(cell.start, cell.startTagEnd);
  const offset = cell.start;
  const removals = MERGE_ATTRIBUTE_NAMES.flatMap((name) =>
    exactAttributes(cell, name).map((attribute) =>
      removeAttributeEdit(source, attribute, offset)));
  const stripped = applyLocalEdits(source, removals);
  const additions = [
    state.rowSpan > 1 ? ` rowSpan="${state.rowSpan}"` : '',
    state.gridSpan > 1 ? ` gridSpan="${state.gridSpan}"` : '',
    state.verticalContinuation ? ' vMerge="1"' : '',
    state.horizontalContinuation ? ' hMerge="1"' : '',
  ].join('');
  const next = additions.length === 0
    ? stripped
    : insertAttributes(stripped, additions);
  if (next !== source) xml.replace(cell.start, cell.startTagEnd, next);
}

function removeAttributeEdit(
  source: string,
  attribute: XmlAttribute,
  offset: number,
): LocalEdit {
  let start = attribute.start - offset;
  while (start > 0 && /[\t\n\r ]/u.test(source[start - 1] ?? '')) start -= 1;
  return {
    start,
    end: attribute.end - offset,
    replacement: '',
  };
}

function insertAttributes(source: string, additions: string): string {
  let position = source.length - 1;
  while (position > 0 && /[\t\n\r ]/u.test(source[position - 1] ?? '')) position -= 1;
  if (source[position - 1] === '/') position -= 1;
  return source.slice(0, position) + additions + source.slice(position);
}

function applyLocalEdits(source: string, edits: readonly LocalEdit[]): string {
  if (edits.length === 0) return source;
  let output = source;
  const ordered = [...edits].sort(
    (left, right) => right.start - left.start || right.end - left.end,
  );
  let previousStart = source.length;
  for (const edit of ordered) {
    if (
      edit.start < 0
      || edit.end < edit.start
      || edit.end > source.length
      || edit.end > previousStart
    ) throw new Error('Overlapping local table merge edits');
    output = output.slice(0, edit.start) + edit.replacement + output.slice(edit.end);
    previousStart = edit.start;
  }
  return output;
}

function exactDirectChild(
  element: XmlElement,
  localName: string,
): XmlElement | undefined {
  const matches = directChildren(element, localName);
  return matches.length === 1 ? matches[0] : undefined;
}

function directChildren(element: XmlElement, localName: string): XmlElement[] {
  return element.children.filter(
    (child): child is XmlElement => child.type === 'element' && child.localName === localName,
  );
}
