import {
  normalizeRichText,
  resolveParagraphSpacing,
  type NormalizedRichTextParagraph,
} from './rich-text.internal.js';
import type {
  NormalizedTableCell,
  NormalizedTableDefinition,
} from './table-create.internal.js';

const DEFAULT_TABLE_FONT_SIZE_PT = 12;
const BASE_CHAR_DIVISOR = 2.3;
const BASE_LINE_MODIFIER = 1.67;
const EMU_PER_POINT = 12_700;
const EMU_PER_INCH = 914_400;
const ZWJ = 0x200D;

interface TextCluster {
  readonly text: string;
  readonly codePoints: readonly number[];
}

interface ResolvedTableCellMargins {
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
  readonly left: number;
}

interface ClusterPiece {
  readonly kind: 'cluster';
  readonly runIndex: number;
  readonly clusterIndex: number;
  readonly clusterCount: number;
  readonly text: string;
  readonly units: number;
  readonly fontSize: number;
  readonly characterSpacing: number;
}

interface EmptyRunPiece {
  readonly kind: 'empty';
  readonly runIndex: number;
}

type LinePiece = ClusterPiece | EmptyRunPiece;

export interface MeasuredTableRunSlice {
  readonly paragraphIndex: number;
  readonly runIndex: number;
  readonly text: string;
  readonly startsAtRunStart: boolean;
  readonly endsAtRunEnd: boolean;
  readonly retainsSoftBreak: boolean;
}

export interface MeasuredTableLine {
  readonly paragraphIndex: number;
  readonly slices: readonly Readonly<MeasuredTableRunSlice>[];
  readonly height: number;
  readonly startsParagraph: boolean;
  readonly endsParagraph: boolean;
}

export interface MeasuredTableCellContent {
  readonly paragraphs: readonly Readonly<NormalizedRichTextParagraph>[];
  readonly lines: readonly Readonly<MeasuredTableLine>[];
  readonly topMargin: number;
  readonly rightMargin: number;
  readonly bottomMargin: number;
  readonly leftMargin: number;
}

export interface MeasuredTableRowLayout {
  readonly sourceRowIndex: number;
  readonly bands: readonly number[];
  readonly contentHeight: number;
  readonly height: number;
  readonly fragmentable: boolean;
  readonly cells: readonly (Readonly<MeasuredTableCellContent> | undefined)[];
}

export function measureTableCellContent(
  cell: Readonly<NormalizedTableCell>,
  width: number,
  tableCharWeight?: number,
  tableLineWeight?: number,
): Readonly<MeasuredTableCellContent> {
  const normalizedWidth = positiveSafeInteger(width, 'Table cell measurement width');
  const charWeight = measurementWeight(
    cell.autoPageCharWeight ?? tableCharWeight ?? 0,
    'Table cell character weight',
  );
  const lineWeight = measurementWeight(
    cell.autoPageLineWeight ?? tableLineWeight ?? 0,
    'Table cell line weight',
  );
  const margins = resolveCellMargins(cell);
  const sourceParagraphs = cell.richText ?? normalizeRichText([{
    runs: [{ text: cell.text, style: {} }],
  }]);
  const paragraphs = detachedFrozen(sourceParagraphs);
  const lines = Object.freeze(paragraphs.flatMap((paragraph, paragraphIndex) =>
    measureParagraph(
      paragraph,
      paragraphIndex,
      cell,
      normalizedWidth,
      margins,
      charWeight,
      lineWeight,
    )));

  return Object.freeze({
    paragraphs,
    lines,
    topMargin: margins.top,
    rightMargin: margins.right,
    bottomMargin: margins.bottom,
    leftMargin: margins.left,
  });
}

export function measureTableAutoPageRows(
  definition: Readonly<NormalizedTableDefinition>,
): readonly Readonly<MeasuredTableRowLayout>[] {
  if (definition.rowHeights.length !== definition.rows.length) {
    throw new RangeError('Measured table row heights must match the row count');
  }
  const charWeight = definition.autoPage?.charWeight;
  const lineWeight = definition.autoPage?.lineWeight;
  return Object.freeze(definition.rows.map((row, rowIndex) => {
    if (row.length !== definition.columnWidths.length) {
      throw new RangeError(`Measured table row ${rowIndex} must match the column count`);
    }
    const cells = row.map((anchor, columnIndex) => {
      if (anchor.continuation !== undefined) return undefined;
      const width = columnSpanWidth(
        definition.columnWidths,
        columnIndex,
        anchor.colspan ?? 1,
      );
      return measureTableCellContent(anchor, width, charWeight, lineWeight);
    });
    const ordinaryCells = row.flatMap((anchor, columnIndex) =>
      anchor.continuation === undefined && anchor.rowspan === undefined
        ? [cells[columnIndex]!]
        : []);
    const bandCount = ordinaryCells.reduce(
      (maximum, measured) => Math.max(maximum, measured.lines.length),
      0,
    );
    const bands = Object.freeze(Array.from({ length: bandCount }, (_, bandIndex) =>
      ordinaryCells.reduce(
        (maximum, measured) => Math.max(
          maximum,
          measured.lines[bandIndex]?.height ?? 0,
        ),
        0,
      )));
    const topMargin = ordinaryCells.reduce(
      (maximum, measured) => Math.max(maximum, measured.topMargin),
      ordinaryCells[0]?.topMargin ?? 0,
    );
    const bottomMargin = ordinaryCells.reduce(
      (maximum, measured) => Math.max(maximum, measured.bottomMargin),
      ordinaryCells[0]?.bottomMargin ?? 0,
    );
    const contentHeight = safeSum(
      [topMargin, ...bands, bottomMargin],
      `Measured table row ${rowIndex} content height`,
    );
    const minimum = nonNegativeSafeInteger(
      definition.rowHeights[rowIndex]!,
      `Measured table row ${rowIndex} minimum`,
    );
    const height = Math.max(1, minimum, contentHeight);
    const fragmentable = ordinaryCells.length > 0 && row.every((candidate) =>
      candidate.rowspan === undefined && candidate.continuation?.vertical !== true);
    return Object.freeze({
      sourceRowIndex: rowIndex,
      bands,
      contentHeight,
      height,
      fragmentable,
      cells: Object.freeze(cells),
    });
  }));
}

export function materializeMeasuredTableRows(
  definition: Readonly<NormalizedTableDefinition>,
  measuredRows: readonly Readonly<MeasuredTableRowLayout>[],
): Readonly<NormalizedTableDefinition> {
  if (definition.autoPage?.measureContent !== true) return definition;
  if (measuredRows.length !== definition.rows.length) {
    throw new RangeError('Measured table rows must match the source row count');
  }

  const heights = measuredRows.map((measured, rowIndex) => {
    if (measured.sourceRowIndex !== rowIndex) {
      throw new RangeError(`Measured table row ${rowIndex} source index is inconsistent`);
    }
    if (measured.cells.length !== definition.rows[rowIndex]!.length) {
      throw new RangeError(`Measured table row ${rowIndex} cells must match the source row`);
    }
    return positiveSafeInteger(
      measured.height,
      `Measured table row ${rowIndex} height`,
    );
  });
  const constraints: Array<{
    start: number;
    span: number;
    column: number;
    required: number;
  }> = [];
  for (const [rowIndex, row] of definition.rows.entries()) {
    for (const [columnIndex, anchor] of row.entries()) {
      if (anchor.continuation !== undefined || anchor.rowspan === undefined) continue;
      const measured = measuredRows[rowIndex]!.cells[columnIndex];
      if (measured === undefined) {
        throw new RangeError(`Measured rowspan anchor ${rowIndex},${columnIndex} is missing`);
      }
      constraints.push({
        start: rowIndex,
        span: anchor.rowspan,
        column: columnIndex,
        required: measuredCellContentHeight(measured, rowIndex, columnIndex),
      });
    }
  }
  constraints.sort((left, right) =>
    left.span - right.span
    || left.start - right.start
    || left.column - right.column);
  for (const constraint of constraints) {
    const end = constraint.start + constraint.span;
    if (end > heights.length) {
      throw new RangeError('Measured rowspan constraint is out of range');
    }
    const current = safeSum(
      heights.slice(constraint.start, end),
      'Measured rowspan current height',
    );
    if (current >= constraint.required) continue;
    const deficit = safeSubtract(
      constraint.required,
      current,
      'Measured rowspan deficit',
    );
    const quotient = Math.floor(deficit / constraint.span);
    const remainder = deficit % constraint.span;
    for (let offset = 0; offset < constraint.span; offset += 1) {
      const rowIndex = constraint.start + offset;
      heights[rowIndex] = safeAdd(
        heights[rowIndex]!,
        quotient + (offset < remainder ? 1 : 0),
        `Measured rowspan row ${rowIndex} height`,
      );
    }
  }

  const rowHeights = Object.freeze(heights);
  const height = safeSum(rowHeights, 'Measured table height');
  const rows = Object.freeze(definition.rows.map((row) => Object.freeze([...row])));
  const columnWidths = Object.freeze([...definition.columnWidths]);
  const autoPage = Object.freeze({
    ...definition.autoPage,
    measureContent: false,
  });
  return Object.freeze({
    ...definition,
    rows,
    height,
    autoRowHeight: false,
    columnWidths,
    rowHeights,
    autoPage,
  });
}

function measureParagraph(
  paragraph: Readonly<NormalizedRichTextParagraph>,
  paragraphIndex: number,
  cell: Readonly<NormalizedTableCell>,
  width: number,
  margins: Readonly<ResolvedTableCellMargins>,
  charWeight: number,
  lineWeight: number,
): readonly Readonly<MeasuredTableLine>[] {
  const pieceLines: LinePiece[][] = [];
  let currentLine: LinePiece[] = [];
  let currentPosition = 0;
  let token: LinePiece[] = [];
  let tokenWhitespace: boolean | undefined;

  const firstLineCapacity = availableLineWidth(
    paragraph,
    width,
    margins,
    true,
  );
  const continuationLineCapacity = availableLineWidth(
    paragraph,
    width,
    margins,
    false,
  );
  const lineCapacity = (): number => pieceLines.length === 0
    ? firstLineCapacity
    : continuationLineCapacity;
  const hasCluster = (): boolean => currentLine.some(({ kind }) => kind === 'cluster');
  const flushLine = (force: boolean): void => {
    if (force || currentLine.length > 0) pieceLines.push(currentLine);
    currentLine = [];
    currentPosition = 0;
  };
  const appendPieces = (pieces: readonly LinePiece[]): void => {
    currentPosition = endPosition(
      pieces,
      currentPosition,
      paragraph,
      charWeight,
    );
    currentLine.push(...pieces);
  };
  const placeToken = (): void => {
    if (token.length === 0) return;
    const capacity = lineCapacity();
    const tokenEnd = endPosition(token, currentPosition, paragraph, charWeight);
    if (tokenEnd <= capacity) {
      appendPieces(token);
      token = [];
      tokenWhitespace = undefined;
      return;
    }

    if (hasCluster()) flushLine(false);
    const emptyCapacity = lineCapacity();
    const emptyEnd = endPosition(token, currentPosition, paragraph, charWeight);
    if (emptyEnd <= emptyCapacity) {
      appendPieces(token);
      token = [];
      tokenWhitespace = undefined;
      return;
    }

    for (const piece of token) {
      if (piece.kind === 'empty') {
        currentLine.push(piece);
        continue;
      }
      const next = endPosition([piece], currentPosition, paragraph, charWeight);
      if (hasCluster() && next > lineCapacity()) flushLine(false);
      appendPieces([piece]);
    }
    token = [];
    tokenWhitespace = undefined;
  };

  for (const [runIndex, run] of paragraph.runs.entries()) {
    if (run.softBreakBefore) {
      placeToken();
      flushLine(true);
    }
    const clusters = textClusters(run.text);
    const fontSize = run.style?.fontSize ?? cell.fontSize ?? DEFAULT_TABLE_FONT_SIZE_PT;
    if (clusters.length === 0) {
      token.push(Object.freeze({ kind: 'empty', runIndex }));
      continue;
    }
    for (const [clusterIndex, cluster] of clusters.entries()) {
      const whitespace = isWhitespaceCluster(cluster);
      if (tokenWhitespace !== undefined && tokenWhitespace !== whitespace) placeToken();
      tokenWhitespace = whitespace;
      token.push(Object.freeze({
        kind: 'cluster',
        runIndex,
        clusterIndex,
        clusterCount: clusters.length,
        text: cluster.text,
        units: clusterUnits(cluster),
        fontSize,
        characterSpacing: run.style?.characterSpacing ?? 0,
      }));
    }
  }
  placeToken();
  flushLine(true);

  const spacing = resolveParagraphSpacing(cell.spacing, paragraph.spacing);
  return Object.freeze(pieceLines.map((pieces, lineIndex) => {
    const startsParagraph = lineIndex === 0;
    const endsParagraph = lineIndex === pieceLines.length - 1;
    const naturalHeight = lineNaturalHeight(
      pieces,
      paragraph,
      cell.fontSize ?? DEFAULT_TABLE_FONT_SIZE_PT,
      lineWeight,
    );
    const spacedHeight = spacing?.line === undefined
      ? naturalHeight
      : spacing.line.kind === 'exact'
        ? pointsToEmu(spacing.line.points, 'Table paragraph exact line spacing')
        : roundedSafeProduct(
            naturalHeight,
            spacing.line.factor,
            'Table paragraph multiple line spacing',
          );
    const before = startsParagraph && spacing?.before !== undefined
      ? pointsToEmu(spacing.before, 'Table paragraph before spacing')
      : 0;
    const after = endsParagraph && spacing?.after !== undefined
      ? pointsToEmu(spacing.after, 'Table paragraph after spacing')
      : 0;
    const height = safeSum(
      [spacedHeight, before, after],
      'Table measured line height',
    );
    return Object.freeze({
      paragraphIndex,
      slices: lineSlices(paragraph, paragraphIndex, pieces),
      height,
      startsParagraph,
      endsParagraph,
    });
  }));
}

function lineSlices(
  paragraph: Readonly<NormalizedRichTextParagraph>,
  paragraphIndex: number,
  pieces: readonly LinePiece[],
): readonly Readonly<MeasuredTableRunSlice>[] {
  const groups: Array<{
    runIndex: number;
    text: string;
    firstCluster?: number;
    lastCluster?: number;
    clusterCount: number;
  }> = [];
  for (const piece of pieces) {
    const previous = groups.at(-1);
    if (previous?.runIndex === piece.runIndex) {
      previous.text += piece.kind === 'cluster' ? piece.text : '';
      if (piece.kind === 'cluster') {
        previous.firstCluster ??= piece.clusterIndex;
        previous.lastCluster = piece.clusterIndex;
        previous.clusterCount = piece.clusterCount;
      }
      continue;
    }
    groups.push({
      runIndex: piece.runIndex,
      text: piece.kind === 'cluster' ? piece.text : '',
      ...(piece.kind === 'cluster'
        ? { firstCluster: piece.clusterIndex, lastCluster: piece.clusterIndex }
        : {}),
      clusterCount: piece.kind === 'cluster' ? piece.clusterCount : 0,
    });
  }
  return Object.freeze(groups.map((group) => {
    const startsAtRunStart = group.clusterCount === 0 || group.firstCluster === 0;
    const endsAtRunEnd = group.clusterCount === 0
      || group.lastCluster === group.clusterCount - 1;
    return Object.freeze({
      paragraphIndex,
      runIndex: group.runIndex,
      text: group.text,
      startsAtRunStart,
      endsAtRunEnd,
      retainsSoftBreak: startsAtRunStart
        && paragraph.runs[group.runIndex]?.softBreakBefore === true,
    });
  }));
}

function lineNaturalHeight(
  pieces: readonly LinePiece[],
  paragraph: Readonly<NormalizedRichTextParagraph>,
  fallbackFontSize: number,
  lineWeight: number,
): number {
  if (pieces.length === 0) return naturalLineHeight(fallbackFontSize, lineWeight);
  let height = 0;
  const runIndexes = new Set(pieces.map(({ runIndex }) => runIndex));
  for (const runIndex of runIndexes) {
    const fontSize = paragraph.runs[runIndex]?.style?.fontSize ?? fallbackFontSize;
    height = Math.max(height, naturalLineHeight(fontSize, lineWeight));
  }
  return height || naturalLineHeight(fallbackFontSize, lineWeight);
}

function endPosition(
  pieces: readonly LinePiece[],
  start: number,
  paragraph: Readonly<NormalizedRichTextParagraph>,
  charWeight: number,
): number {
  let position = start;
  for (const piece of pieces) {
    if (piece.kind === 'empty') continue;
    const advance = piece.text.codePointAt(0) === 0x09
      ? tabAdvance(position, piece, paragraph, charWeight)
      : clusterAdvance(piece, charWeight);
    position = safeAdd(position, advance, 'Table measured inline position');
  }
  return position;
}

function tabAdvance(
  position: number,
  piece: Readonly<ClusterPiece>,
  paragraph: Readonly<NormalizedRichTextParagraph>,
  charWeight: number,
): number {
  const declaredStops = paragraph.tabStops === false
    ? []
    : paragraph.tabStops ?? [];
  const target = declaredStops.reduce<number | undefined>((nearest, stop) =>
    stop.positionEmu > position && (nearest === undefined || stop.positionEmu < nearest)
      ? stop.positionEmu
      : nearest, undefined);
  if (target !== undefined) return target - position;

  const spaceAdvance = measuredAdvance(
    piece.fontSize,
    0.5,
    piece.characterSpacing,
    charWeight,
  );
  const grid = safeProduct(spaceAdvance, 4, 'Table tab grid');
  const gridIndex = Math.floor(position / grid) + 1;
  const next = safeProduct(grid, gridIndex, 'Table tab position');
  return next - position;
}

function clusterAdvance(piece: Readonly<ClusterPiece>, charWeight: number): number {
  return measuredAdvance(
    piece.fontSize,
    piece.units,
    piece.characterSpacing,
    charWeight,
  );
}

function measuredAdvance(
  fontSize: number,
  units: number,
  characterSpacing: number,
  charWeight: number,
): number {
  const value = fontSize * EMU_PER_POINT * units / (BASE_CHAR_DIVISOR + charWeight)
    + characterSpacing * EMU_PER_POINT;
  if (!Number.isFinite(value)) {
    throw new RangeError('Table character advance must be finite');
  }
  const rounded = Math.round(value);
  if (!Number.isSafeInteger(rounded)) {
    throw new RangeError('Table character advance must fit a safe integer EMU value');
  }
  return Math.max(1, rounded);
}

function textClusters(value: string): readonly Readonly<TextCluster>[] {
  const clusters: Array<{ text: string; codePoints: number[] }> = [];
  for (const character of Array.from(value)) {
    const codePoint = character.codePointAt(0)!;
    const previous = clusters.at(-1);
    if (
      previous
      && (isClusterAttachment(codePoint)
        || codePoint === ZWJ
        || previous.codePoints.at(-1) === ZWJ)
    ) {
      previous.text += character;
      previous.codePoints.push(codePoint);
    } else {
      clusters.push({ text: character, codePoints: [codePoint] });
    }
  }
  return Object.freeze(clusters.map((cluster) => Object.freeze({
    text: cluster.text,
    codePoints: Object.freeze(cluster.codePoints),
  })));
}

function clusterUnits(cluster: Readonly<TextCluster>): number {
  const base = cluster.codePoints[0] ?? 0;
  if (isEcmaWhitespace(base)) return 0.5;
  if (
    (base >= 0x21 && base <= 0x2F)
    || (base >= 0x3A && base <= 0x40)
    || (base >= 0x5B && base <= 0x60)
    || (base >= 0x7B && base <= 0x7E)
  ) return 0.6;
  if (
    (base >= 0x30 && base <= 0x39)
    || (base >= 0x41 && base <= 0x5A)
    || (base >= 0x61 && base <= 0x7A)
    || (base >= 0xC0 && base <= 0x24F)
    || (base >= 0x370 && base <= 0x3FF)
    || (base >= 0x400 && base <= 0x52F)
  ) return 1;
  return 2.3;
}

function isWhitespaceCluster(cluster: Readonly<TextCluster>): boolean {
  return isEcmaWhitespace(cluster.codePoints[0] ?? 0);
}

function isEcmaWhitespace(codePoint: number): boolean {
  return /\s/u.test(String.fromCodePoint(codePoint));
}

function isClusterAttachment(codePoint: number): boolean {
  return (codePoint >= 0x300 && codePoint <= 0x36F)
    || (codePoint >= 0x1AB0 && codePoint <= 0x1AFF)
    || (codePoint >= 0x1DC0 && codePoint <= 0x1DFF)
    || (codePoint >= 0x20D0 && codePoint <= 0x20FF)
    || (codePoint >= 0xFE20 && codePoint <= 0xFE2F)
    || (codePoint >= 0xFE00 && codePoint <= 0xFE0F)
    || (codePoint >= 0xE0100 && codePoint <= 0xE01EF)
    || (codePoint >= 0x1F3FB && codePoint <= 0x1F3FF);
}

function availableLineWidth(
  paragraph: Readonly<NormalizedRichTextParagraph>,
  width: number,
  margins: Readonly<ResolvedTableCellMargins>,
  firstLine: boolean,
): number {
  const marginLeft = typeof paragraph.marginLeft === 'number'
    ? pointsToEmu(paragraph.marginLeft, 'Table paragraph left margin')
    : 0;
  const indent = firstLine && typeof paragraph.indent === 'number'
    ? pointsToEmu(paragraph.indent, 'Table paragraph first-line indent')
    : 0;
  const ordinaryLeft = firstLine
    ? Math.max(0, safeAdd(marginLeft, indent, 'Table paragraph first-line reserve'))
    : Math.max(0, marginLeft);
  const bulletLeft = firstLine
    && paragraph.bullet !== undefined
    && paragraph.bullet !== false
    ? pointsToEmu(paragraph.bullet.indent, 'Table paragraph bullet indent')
    : 0;
  const leftReserve = Math.max(ordinaryLeft, bulletLeft);
  const rightReserve = typeof paragraph.marginRight === 'number'
    ? Math.max(0, pointsToEmu(paragraph.marginRight, 'Table paragraph right margin'))
    : 0;
  let available = safeSubtract(width, margins.left, 'Table cell inline width');
  available = safeSubtract(available, margins.right, 'Table cell inline width');
  available = safeSubtract(available, leftReserve, 'Table cell inline width');
  available = safeSubtract(available, rightReserve, 'Table cell inline width');
  if (available <= 0) {
    throw new RangeError('Table cell usable inline width must be positive');
  }
  return available;
}

function naturalLineHeight(fontSize: number, lineWeight: number): number {
  const value = fontSize * (BASE_LINE_MODIFIER + lineWeight) * EMU_PER_INCH / 100;
  const rounded = Math.round(value);
  if (!Number.isSafeInteger(rounded) || rounded <= 0) {
    throw new RangeError('Table natural line height must be a positive safe integer EMU value');
  }
  return rounded;
}

function resolveCellMargins(
  cell: Readonly<NormalizedTableCell>,
): Readonly<ResolvedTableCellMargins> {
  return Object.freeze({
    top: pointsToEmu(cell.margins?.top ?? 3.6, 'Table cell top margin'),
    right: pointsToEmu(cell.margins?.right ?? 7.2, 'Table cell right margin'),
    bottom: pointsToEmu(cell.margins?.bottom ?? 3.6, 'Table cell bottom margin'),
    left: pointsToEmu(cell.margins?.left ?? 7.2, 'Table cell left margin'),
  });
}

function columnSpanWidth(
  columnWidths: readonly number[],
  start: number,
  span: number,
): number {
  if (!Number.isSafeInteger(span) || span < 1 || start + span > columnWidths.length) {
    throw new RangeError('Measured table cell colspan is out of range');
  }
  return safeSum(
    columnWidths.slice(start, start + span).map((width, offset) =>
      positiveSafeInteger(width, `Measured table column ${start + offset} width`)),
    'Measured table cell colspan width',
  );
}

function measuredCellContentHeight(
  measured: Readonly<MeasuredTableCellContent>,
  rowIndex: number,
  columnIndex: number,
): number {
  return safeSum([
    measured.topMargin,
    ...measured.lines.map(({ height }, lineIndex) => positiveSafeInteger(
      height,
      `Measured rowspan ${rowIndex},${columnIndex} line ${lineIndex} height`,
    )),
    measured.bottomMargin,
  ], `Measured rowspan ${rowIndex},${columnIndex} content height`);
}

function measurementWeight(value: number, context: string): number {
  if (!Number.isFinite(value)) throw new TypeError(`${context} must be finite`);
  if (value < -1 || value > 1) {
    throw new RangeError(`${context} must be between -1 and 1`);
  }
  return Object.is(value, -0) ? 0 : value;
}

function pointsToEmu(value: number, context: string): number {
  if (!Number.isFinite(value)) throw new TypeError(`${context} must be finite`);
  const result = Math.round(value * EMU_PER_POINT);
  if (!Number.isSafeInteger(result)) {
    throw new RangeError(`${context} must fit a safe integer EMU value`);
  }
  return result;
}

function positiveSafeInteger(value: number, context: string): number {
  if (!Number.isFinite(value)) throw new TypeError(`${context} must be finite`);
  if (!Number.isSafeInteger(value)) {
    throw new TypeError(`${context} must be a safe integer EMU value`);
  }
  if (value <= 0) throw new RangeError(`${context} must be positive`);
  return value;
}

function nonNegativeSafeInteger(value: number, context: string): number {
  if (!Number.isFinite(value)) throw new TypeError(`${context} must be finite`);
  if (!Number.isSafeInteger(value)) {
    throw new TypeError(`${context} must be a safe integer EMU value`);
  }
  if (value < 0) throw new RangeError(`${context} must be non-negative`);
  return value;
}

function safeAdd(left: number, right: number, context: string): number {
  const result = left + right;
  if (!Number.isSafeInteger(result)) {
    throw new RangeError(`${context} must fit a safe integer EMU value`);
  }
  return result;
}

function safeSubtract(left: number, right: number, context: string): number {
  const result = left - right;
  if (!Number.isSafeInteger(result)) {
    throw new RangeError(`${context} must fit a safe integer EMU value`);
  }
  return result;
}

function safeSum(values: readonly number[], context: string): number {
  return values.reduce((sum, value) => safeAdd(sum, value, context), 0);
}

function safeProduct(left: number, right: number, context: string): number {
  const result = left * right;
  if (!Number.isSafeInteger(result)) {
    throw new RangeError(`${context} must fit a safe integer EMU value`);
  }
  return result;
}

function roundedSafeProduct(left: number, right: number, context: string): number {
  const result = Math.round(left * right);
  if (!Number.isSafeInteger(result) || result <= 0) {
    throw new RangeError(`${context} must be a positive safe integer EMU value`);
  }
  return result;
}

function detachedFrozen<T>(value: T): T {
  if (Array.isArray(value)) {
    return Object.freeze(value.map((item) => detachedFrozen(item))) as T;
  }
  if (value && typeof value === 'object') {
    const result: Record<string, unknown> = {};
    for (const [key, item] of Object.entries(value)) result[key] = detachedFrozen(item);
    return Object.freeze(result) as T;
  }
  return value;
}
