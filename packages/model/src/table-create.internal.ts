import { escapeXmlAttribute } from '@pptx/lossless-xml';
import { normalizePlaceholderSelector } from './placeholder.internal.js';
import type { PlaceholderIdentity, PlaceholderSelector } from './placeholder.js';
import type { Transform } from './units.js';
import {
  normalizeRichText,
  normalizeTextAlignment,
  renderRichTextParagraphs,
  type RichTextRunHyperlinkRelationshipIds,
} from './rich-text.internal.js';
import {
  normalizeHyperlink,
  type NormalizedHyperlink,
} from './shape-hyperlink.internal.js';
import {
  normalizeTableCellFill,
  renderTableCellFill,
} from './table-cell-fill.internal.js';
import {
  normalizeTableCellBorders,
  renderTableCellBorders,
} from './table-cell-borders.internal.js';
import { renderTableCellMarginAttributes } from './table-cell-margins.internal.js';
import {
  normalizeTableCellTextDirection,
  renderTableCellTextDirectionAttribute,
} from './table-cell-text-direction.internal.js';
import {
  renderTableCellVerticalAlignmentAttribute,
} from './table-cell-vertical-alignment.internal.js';
import {
  normalizeTextBoxFit,
  renderTextBoxFitChild,
} from './text-box-fit.internal.js';
import { normalizeTextBoxMargins } from './text-box-margins.internal.js';
import {
  normalizeTextBoxVerticalAlignment,
} from './text-box-vertical-alignment.internal.js';
import type {
  TableCellBorders,
  TableCellFill,
  TableCellTextDirection,
} from './shapes.js';
import type {
  TextAlignment,
  TextBoxFit,
  TextBoxMarginInput,
  TextBoxMargins,
  TextBoxVerticalAlignment,
} from './text.js';

const EMU_PER_INCH = 914_400;
const DEFAULT_OFFSET = EMU_PER_INCH / 2;
const DEFAULT_HEIGHT = EMU_PER_INCH;
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OPTION_KEYS = [
  'name',
  'placeholder',
  'x',
  'y',
  'width',
  'height',
  'columnWidths',
  'rowHeights',
  'align',
  'border',
  'fill',
  'margin',
  'textDirection',
  'valign',
] as const;
interface NormalizedTableCell {
  readonly text: string;
  readonly richText?: ReturnType<typeof normalizeRichText>;
  readonly alignment?: TextAlignment;
  readonly borders?: TableCellBorders;
  readonly fill?: TableCellFill;
  readonly hyperlink?: NormalizedHyperlink;
  readonly margins?: TextBoxMargins;
  readonly textDirection?: TableCellTextDirection;
  readonly textFit?: TextBoxFit;
  readonly verticalAlignment?: TextBoxVerticalAlignment;
}

export interface NormalizedTableDefinition {
  readonly rows: readonly (readonly NormalizedTableCell[])[];
  readonly name?: string;
  readonly placeholder?: PlaceholderSelector;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly autoRowHeight: boolean;
  readonly columnWidths: readonly number[];
  readonly rowHeights: readonly number[];
}

export type TableCellHyperlinkRelationshipIds =
  readonly (readonly (string | undefined)[])[];

export type TableCellRichTextRunHyperlinkRelationshipIds =
  readonly (readonly RichTextRunHyperlinkRelationshipIds[])[];

export function normalizeTableDefinition(
  rows: unknown,
  options: unknown,
): NormalizedTableDefinition {
  const outer = readDenseArray(rows, 'Table rows');
  const normalizedRows = outer.map((row, rowIndex) =>
    readDenseArray(row, `Table row ${rowIndex}`).map((cell, columnIndex) =>
      normalizeTableCell(cell, rowIndex, columnIndex)));
  const columnCount = normalizedRows[0]!.length;
  for (let rowIndex = 1; rowIndex < normalizedRows.length; rowIndex += 1) {
    if (normalizedRows[rowIndex]!.length !== columnCount) {
      throw new TypeError('Table rows must contain the same number of cells');
    }
  }

  const normalizedOptions = readOptions(options);
  const placeholder = normalizedOptions.placeholder === undefined
    ? undefined
    : normalizePlaceholderSelector(normalizedOptions.placeholder);
  const tableAlignment = normalizedOptions.align === undefined
    ? undefined
    : normalizeTextAlignment(normalizedOptions.align, 'Table align');
  const alignmentResolvedRows = tableAlignment === undefined
    ? normalizedRows
    : normalizedRows.map((row) => row.map((cell) =>
      cell.alignment === undefined
        ? { ...cell, alignment: tableAlignment }
        : cell));
  const tableBorders = normalizeTableCellBorders(
    normalizedOptions.border,
    'Table border',
  );
  const borderResolvedRows = tableBorders === undefined
    ? alignmentResolvedRows
    : alignmentResolvedRows.map((row) => row.map((cell) =>
      cell.borders === undefined
        ? { ...cell, borders: tableBorders }
        : cell));
  const tableFill = normalizeTableCellFill(normalizedOptions.fill, 'Table fill');
  const fillResolvedRows = tableFill === undefined
    ? borderResolvedRows
    : borderResolvedRows.map((row) => row.map((cell) =>
      cell.fill === undefined
        ? { ...cell, fill: tableFill }
        : cell));
  const tableMargins = normalizeTextBoxMargins(
    normalizedOptions.margin as TextBoxMarginInput | undefined,
    'Table margin',
  );
  const marginResolvedRows = tableMargins === undefined
    ? fillResolvedRows
    : fillResolvedRows.map((row) => row.map((cell) => ({
      ...cell,
      margins: { ...tableMargins, ...(cell.margins ?? {}) },
    })));
  const tableTextDirection = normalizedOptions.textDirection === undefined
    ? undefined
    : normalizeTableCellTextDirection(
      normalizedOptions.textDirection,
      'Table textDirection',
    );
  const directionResolvedRows = tableTextDirection === undefined
    ? marginResolvedRows
    : marginResolvedRows.map((row) => row.map((cell) =>
      cell.textDirection === undefined
        ? { ...cell, textDirection: tableTextDirection }
        : cell));
  const tableVerticalAlignment = normalizedOptions.valign === undefined
    ? undefined
    : normalizeTextBoxVerticalAlignment(normalizedOptions.valign, 'Table valign');
  const resolvedRows = tableVerticalAlignment === undefined
    ? directionResolvedRows
    : directionResolvedRows.map((row) => row.map((cell) =>
      cell.verticalAlignment === undefined
        ? { ...cell, verticalAlignment: tableVerticalAlignment }
        : cell));
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
    rows: resolvedRows,
    ...(name !== undefined ? { name } : {}),
    ...(placeholder === undefined ? {} : { placeholder }),
    x,
    y,
    width,
    height,
    autoRowHeight,
    columnWidths,
    rowHeights,
  };
}

function normalizeTableCell(
  cell: unknown,
  rowIndex: number,
  columnIndex: number,
): NormalizedTableCell {
  const context = `Table cell ${rowIndex},${columnIndex}`;
  if (typeof cell === 'string') {
    return normalizeTableCellText(cell, context);
  }
  if (!cell || typeof cell !== 'object' || Array.isArray(cell)) {
    throw new TypeError(`${context} must be a string or text object`);
  }
  const candidate = readDataObject(cell, context, ['text', 'options']);
  const text = normalizeTableCellText(candidate.text, context);
  return {
    ...text,
    ...normalizeTableCellOptions(candidate.options, context),
  };
}

function normalizeTableCellText(
  value: unknown,
  context: string,
): Pick<NormalizedTableCell, 'text' | 'richText'> {
  if (typeof value === 'string') {
    if (containsInvalidXmlCharacter(value)) {
      throw new TypeError(`${context} contains invalid XML characters`);
    }
    const text = value.replace(/\r\n?/g, '\n');
    if (!text.includes('\n')) return { text };
    return {
      text,
      richText: normalizeRichText(text.split('\n').map((line) => ({
        runs: [{ text: line, style: {} }],
      }))),
    };
  }
  const richText = normalizeRichText(value);
  return {
    text: projectTableCellText(richText),
    richText,
  };
}

function normalizeTableCellOptions(
  value: unknown,
  context: string,
): Pick<
  NormalizedTableCell,
  | 'alignment'
  | 'borders'
  | 'fill'
  | 'hyperlink'
  | 'margins'
  | 'textDirection'
  | 'textFit'
  | 'verticalAlignment'
> {
  if (value === undefined) return {};
  const options = readDataObject(
    value,
    `${context} options`,
    ['align', 'border', 'fill', 'fit', 'hyperlink', 'margin', 'textDirection', 'valign'],
  );
  const alignment = options.align === undefined
    ? undefined
    : normalizeTextAlignment(options.align, `${context} align`);
  const borders = normalizeTableCellBorders(options.border, `${context} border`);
  const fill = normalizeTableCellFill(options.fill, `${context} fill`);
  const hyperlink = options.hyperlink === undefined
    ? undefined
    : normalizeHyperlink(options.hyperlink, `${context} hyperlink`);
  const margins = normalizeTextBoxMargins(
    options.margin as TextBoxMarginInput | undefined,
    `${context} margin`,
  );
  const textDirection = options.textDirection === undefined
    ? undefined
    : normalizeTableCellTextDirection(
      options.textDirection,
      `${context} textDirection`,
    );
  const textFit = options.fit === undefined
    ? undefined
    : normalizeTextBoxFit(options.fit, `${context} fit`);
  const verticalAlignment = options.valign === undefined
    ? undefined
    : normalizeTextBoxVerticalAlignment(options.valign, `${context} valign`);
  return {
    ...(alignment === undefined ? {} : { alignment }),
    ...(borders === undefined ? {} : { borders }),
    ...(fill === undefined ? {} : { fill }),
    ...(hyperlink === undefined ? {} : { hyperlink }),
    ...(margins === undefined ? {} : { margins }),
    ...(textDirection === undefined ? {} : { textDirection }),
    ...(textFit === undefined ? {} : { textFit }),
    ...(verticalAlignment === undefined ? {} : { verticalAlignment }),
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
  placeholder?: Readonly<PlaceholderIdentity>,
  transform?: Readonly<Transform>,
  hyperlinkRelationshipIds?: TableCellHyperlinkRelationshipIds,
  richTextRunHyperlinkRelationshipIds?: TableCellRichTextRunHyperlinkRelationshipIds,
): string {
  const relationshipIds = hyperlinkRelationshipIds ?? definition.rows.map((row) =>
    row.map(() => undefined));
  if (relationshipIds.length !== definition.rows.length) {
    throw new TypeError('Table-cell hyperlink relationship IDs must match the row count');
  }
  const grid = definition.columnWidths.map((width) => `<a:gridCol w="${width}"/>`).join('');
  const rows = definition.rows.map((row, rowIndex) => {
    if (relationshipIds[rowIndex]?.length !== row.length) {
      throw new TypeError(
        `Table-cell hyperlink relationship IDs must match row ${rowIndex} cell count`,
      );
    }
    const runRelationshipIds = richTextRunHyperlinkRelationshipIds?.[rowIndex];
    if (runRelationshipIds !== undefined && runRelationshipIds.length !== row.length) {
      throw new TypeError(
        `Table-cell rich-text hyperlink relationship IDs must match row ${rowIndex} cell count`,
      );
    }
    const cells = row.map((cell, columnIndex) => renderTableCell(
      cell,
      relationshipIds[rowIndex]![columnIndex],
      runRelationshipIds?.[columnIndex],
    )).join('');
    return `<a:tr h="${definition.rowHeights[rowIndex]}">${cells}</a:tr>`;
  }).join('');
  const name = escapeXmlAttribute(definition.name ?? `Table ${id}`);
  const applicationProperties = placeholder === undefined
    ? '<p:nvPr/>'
    : `<p:nvPr><p:ph type="${placeholder.type}" idx="${placeholder.index}"/></p:nvPr>`;
  const rotation = transform?.rotation ?? 0;
  const transformAttributes = [
    rotation === 0 ? '' : ` rot="${rotation}"`,
    transform?.flipHorizontal ? ' flipH="1"' : '',
    transform?.flipVertical ? ' flipV="1"' : '',
  ].join('');
  if (
    richTextRunHyperlinkRelationshipIds !== undefined
    && richTextRunHyperlinkRelationshipIds.length !== definition.rows.length
  ) {
    throw new TypeError(
      'Table-cell rich-text hyperlink relationship IDs must match the row count',
    );
  }
  const relationshipNamespace = definition.rows.some((row) =>
    row.some((cell) => tableCellHasRenderedHyperlink(cell)))
    ? ` xmlns:r="${RELATIONSHIP_NAMESPACE}"`
    : '';

  return `<p:graphicFrame xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"${relationshipNamespace}><p:nvGraphicFramePr><p:cNvPr id="${id}" name="${name}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>${applicationProperties}</p:nvGraphicFramePr><p:xfrm${transformAttributes}><a:off x="${definition.x}" y="${definition.y}"/><a:ext cx="${definition.width}" cy="${definition.height}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/><a:tblGrid>${grid}</a:tblGrid>${rows}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
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

function readDataObject(
  value: unknown,
  context: string,
  supported: readonly string[],
): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${context} must be an object`);
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError(`${context} must be an ordinary object`);
  }
  const allowed = new Set(supported);
  const result: Record<string, unknown> = {};
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !allowed.has(key)) {
      throw new TypeError(`${context} contains unsupported property ${String(key)}`);
    }
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`${context} property ${key} must be a data property`);
    }
    result[key] = descriptor.value;
  }
  return result;
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

function renderTableCell(
  cell: NormalizedTableCell,
  hyperlinkRelationshipId: string | undefined,
  runHyperlinkRelationshipIds?: RichTextRunHyperlinkRelationshipIds,
): string {
  const paragraphs = tableCellRichText(cell);
  const defaultHyperlink = cell.hyperlink !== undefined
    && paragraphs.some(({ runs }) =>
      runs.some(({ style }) => style?.hyperlink === undefined))
    ? cell.hyperlink
    : undefined;
  if ((defaultHyperlink === undefined) !== (hyperlinkRelationshipId === undefined)) {
    throw new TypeError('Table-cell hyperlink and relationship ID must be supplied together');
  }
  const renderedParagraphs = renderRichTextParagraphs(paragraphs, {
    ...(cell.alignment === undefined ? {} : { defaultAlign: cell.alignment }),
    ...(defaultHyperlink === undefined
      ? {}
      : {
          defaultHyperlink,
          hyperlinkRelationshipId: hyperlinkRelationshipId!,
        }),
    ...(runHyperlinkRelationshipIds === undefined
      ? {}
      : { runHyperlinkRelationshipIds }),
  });
  const borders = renderTableCellBorders(cell.borders, 'a:');
  const fill = cell.fill === undefined ? '' : renderTableCellFill(cell.fill, 'a:');
  const marginAttributes = renderTableCellMarginAttributes(cell.margins);
  const verticalAlignmentAttribute = renderTableCellVerticalAlignmentAttribute(
    cell.verticalAlignment,
  );
  const textDirectionAttribute = renderTableCellTextDirectionAttribute(
    cell.textDirection,
  );
  const textFitChild = cell.textFit === undefined
    ? ''
    : renderTextBoxFitChild(cell.textFit);
  const bodyProperties = textFitChild === ''
    ? '<a:bodyPr/>'
    : `<a:bodyPr>${textFitChild}</a:bodyPr>`;
  return `<a:tc><a:txBody>${bodyProperties}<a:lstStyle/>${renderedParagraphs}</a:txBody><a:tcPr${marginAttributes}${verticalAlignmentAttribute}${textDirectionAttribute}>${borders}${fill}</a:tcPr></a:tc>`;
}

function tableCellRichText(
  cell: NormalizedTableCell,
): ReturnType<typeof normalizeRichText> {
  return cell.richText ?? normalizeRichText([
    { runs: [{ text: cell.text, style: {} }] },
  ]);
}

function tableCellHasRenderedHyperlink(cell: NormalizedTableCell): boolean {
  return tableCellRichText(cell).some(({ runs }) => runs.some(({ style }) => {
    const local = style?.hyperlink;
    return local !== false && (local !== undefined || cell.hyperlink !== undefined);
  }));
}

function projectTableCellText(
  paragraphs: ReturnType<typeof normalizeRichText>,
): string {
  return paragraphs.map(({ runs }) => runs.map((run) =>
    `${run.softBreakBefore ? '\n' : ''}${run.text}`).join('')).join('\n');
}

function containsInvalidXmlCharacter(value: string): boolean {
  return /[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(value);
}
