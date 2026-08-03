import { escapeXmlAttribute } from '@pptx/lossless-xml';
import { normalizePlaceholderSelector } from './placeholder.internal.js';
import type { PlaceholderIdentity, PlaceholderSelector } from './placeholder.js';
import type { Transform } from './units.js';
import {
  normalizeParagraphSpacing,
  normalizeRichText,
  normalizeRichTextColor,
  normalizeTextAlignment,
  renderRichTextParagraphs,
  type NormalizedParagraphSpacingUpdate,
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
  RichTextColor,
  TextAlignment,
  TextBoxFit,
  TextBoxMarginInput,
  TextBoxMargins,
  TextBoxVerticalAlignment,
} from './text.js';
import type { TableCellMergeTokens } from './table-cell-merge.internal.js';
import {
  normalizeTableAutoPageRequest,
  type NormalizedTableAutoPageRequest,
} from './table-auto-page.internal.js';

const EMU_PER_INCH = 914_400;
const DEFAULT_OFFSET = EMU_PER_INCH / 2;
const DEFAULT_HEIGHT = EMU_PER_INCH;
const MAX_TABLE_PHYSICAL_CELLS = 1_000_000;
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OPTION_KEYS = [
  'name',
  'placeholder',
  'autoPage',
  'autoPageRepeatHeader',
  'autoPageHeaderRows',
  'autoPageSlideStartY',
  'slideMargin',
  'x',
  'y',
  'width',
  'height',
  'columnWidths',
  'rowHeights',
  'align',
  'bold',
  'border',
  'color',
  'fill',
  'fontFamily',
  'fontSize',
  'margin',
  'spacing',
  'textDirection',
  'valign',
] as const;

interface NormalizedTableTextDefaults {
  readonly fontFamily?: string;
  readonly fontSize?: number;
  readonly bold?: boolean;
  readonly color?: Readonly<RichTextColor>;
  readonly spacing?: NormalizedParagraphSpacingUpdate;
}

export interface NormalizedTableCell extends NormalizedTableTextDefaults {
  readonly text: string;
  readonly richText?: ReturnType<typeof normalizeRichText>;
  readonly colspan?: number;
  readonly rowspan?: number;
  readonly continuation?: Readonly<{
    readonly rowSpan?: number;
    readonly gridSpan?: number;
    readonly vertical?: true;
    readonly horizontal?: true;
  }>;
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
  readonly autoPage?: Readonly<NormalizedTableAutoPageRequest>;
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
    readDenseArray(row, `Table row ${rowIndex}`, rowIndex > 0).map((cell, columnIndex) =>
      normalizeTableCell(cell, rowIndex, columnIndex)));

  const normalizedOptions = readOptions(options);
  const placeholder = normalizedOptions.placeholder === undefined
    ? undefined
    : normalizePlaceholderSelector(normalizedOptions.placeholder);
  const tableTextDefaults = normalizeTableTextDefaults(normalizedOptions, 'Table');
  const textResolvedRows = Object.keys(tableTextDefaults).length === 0
    ? normalizedRows
    : normalizedRows.map((row) => row.map((cell) =>
      resolveTableTextDefaults(tableTextDefaults, cell)));
  const tableAlignment = normalizedOptions.align === undefined
    ? undefined
    : normalizeTextAlignment(normalizedOptions.align, 'Table align');
  const alignmentResolvedRows = tableAlignment === undefined
    ? textResolvedRows
    : textResolvedRows.map((row) => row.map((cell) =>
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
  const physicalRows = expandTableRows(resolvedRows);
  const columnCount = physicalRows[0]!.length;
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
      physicalRows.length,
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
    if (!autoRowHeight && height < physicalRows.length) {
      throw new RangeError('Table height must provide at least one EMU per row');
    }
    rowHeights = autoRowHeight
      ? physicalRows.map(() => 0)
      : distributeTableDimension(height, physicalRows.length);
  }

  const autoPage = normalizeTableAutoPageRequest(normalizedOptions, {
    rowCount: physicalRows.length,
    rowHeights,
    autoRowHeight,
    hasPlaceholder: placeholder !== undefined,
  });

  return {
    rows: physicalRows,
    ...(name !== undefined ? { name } : {}),
    ...(placeholder === undefined ? {} : { placeholder }),
    x,
    y,
    width,
    height,
    autoRowHeight,
    columnWidths,
    rowHeights,
    ...(autoPage === undefined ? {} : { autoPage }),
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
  | 'bold'
  | 'borders'
  | 'color'
  | 'colspan'
  | 'fill'
  | 'fontFamily'
  | 'fontSize'
  | 'hyperlink'
  | 'margins'
  | 'rowspan'
  | 'spacing'
  | 'textDirection'
  | 'textFit'
  | 'verticalAlignment'
> {
  if (value === undefined) return {};
  const options = readDataObject(
    value,
    `${context} options`,
    [
      'align',
      'bold',
      'border',
      'color',
      'colspan',
      'fill',
      'fit',
      'fontFamily',
      'fontSize',
      'hyperlink',
      'margin',
      'rowspan',
      'spacing',
      'textDirection',
      'valign',
    ],
  );
  const textDefaults = normalizeTableTextDefaults(options, context);
  const colspan = normalizeTableCellSpan(options.colspan, `${context} colspan`);
  const rowspan = normalizeTableCellSpan(options.rowspan, `${context} rowspan`);
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
    ...textDefaults,
    ...(colspan === undefined ? {} : { colspan }),
    ...(alignment === undefined ? {} : { alignment }),
    ...(borders === undefined ? {} : { borders }),
    ...(fill === undefined ? {} : { fill }),
    ...(hyperlink === undefined ? {} : { hyperlink }),
    ...(margins === undefined ? {} : { margins }),
    ...(rowspan === undefined ? {} : { rowspan }),
    ...(textDirection === undefined ? {} : { textDirection }),
    ...(textFit === undefined ? {} : { textFit }),
    ...(verticalAlignment === undefined ? {} : { verticalAlignment }),
  };
}

function normalizeTableCellSpan(
  value: unknown,
  context: string,
): number | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  if (!Number.isSafeInteger(value) || value < 1) {
    throw new RangeError(`${context} must be a positive safe integer`);
  }
  return value === 1 ? undefined : value;
}

function normalizeTableTextDefaults(
  options: Readonly<Record<string, unknown>>,
  context: string,
): NormalizedTableTextDefaults {
  const fontFamily = options.fontFamily;
  if (fontFamily !== undefined) {
    if (typeof fontFamily !== 'string' || fontFamily.length === 0) {
      throw new TypeError(`${context} fontFamily must be a non-empty string`);
    }
    if (containsInvalidXmlCharacter(fontFamily)) {
      throw new TypeError(`${context} fontFamily contains invalid XML characters`);
    }
  }
  const requestedFontSize = options.fontSize;
  if (
    requestedFontSize !== undefined
    && (typeof requestedFontSize !== 'number' || !Number.isFinite(requestedFontSize))
  ) {
    throw new TypeError(`${context} fontSize must be finite`);
  }
  if (
    requestedFontSize !== undefined
    && (requestedFontSize < 1 || requestedFontSize > 4000)
  ) {
    throw new RangeError(`${context} fontSize must be between 1 and 4000 points`);
  }
  const bold = options.bold;
  if (bold !== undefined && typeof bold !== 'boolean') {
    throw new TypeError(`${context} bold must be a boolean`);
  }
  const color = options.color === undefined
    ? undefined
    : normalizeRichTextColor(options.color, `${context} color`);
  const spacing = options.spacing === undefined
    ? undefined
    : normalizeParagraphSpacing(options.spacing, `${context} spacing`);
  return {
    ...(fontFamily === undefined ? {} : { fontFamily }),
    ...(requestedFontSize === undefined
      ? {}
      : { fontSize: Math.round(requestedFontSize * 100) / 100 }),
    ...(bold === undefined ? {} : { bold }),
    ...(color === undefined ? {} : { color }),
    ...(spacing === undefined ? {} : { spacing }),
  };
}

function resolveTableTextDefaults(
  table: NormalizedTableTextDefaults,
  cell: NormalizedTableCell,
): NormalizedTableCell {
  const fontFamily = cell.fontFamily ?? table.fontFamily;
  const fontSize = cell.fontSize ?? table.fontSize;
  const bold = cell.bold ?? table.bold;
  const color = cell.color ?? table.color;
  const spacing = table.spacing === undefined && cell.spacing === undefined
    ? undefined
    : { ...table.spacing, ...cell.spacing };
  return {
    ...cell,
    ...(fontFamily === undefined ? {} : { fontFamily }),
    ...(fontSize === undefined ? {} : { fontSize }),
    ...(bold === undefined ? {} : { bold }),
    ...(color === undefined ? {} : { color }),
    ...(spacing === undefined ? {} : { spacing }),
  };
}

function expandTableRows(
  logicalRows: readonly (readonly NormalizedTableCell[])[],
): readonly (readonly NormalizedTableCell[])[] {
  const columnCount = logicalRows[0]!.reduce((sum, cell) => {
    const colspan = cell.colspan ?? 1;
    if (colspan > Number.MAX_SAFE_INTEGER - sum) {
      throw new RangeError('Table physical column count must fit a safe integer');
    }
    return sum + colspan;
  }, 0);
  if (columnCount > Math.floor(MAX_TABLE_PHYSICAL_CELLS / logicalRows.length)) {
    throw new RangeError(
      `Table cannot contain more than ${MAX_TABLE_PHYSICAL_CELLS} physical cells`,
    );
  }
  const matrix: Array<Array<NormalizedTableCell | undefined>> = logicalRows.map(() =>
    Array.from({ length: columnCount }, () => undefined));

  for (let rowIndex = 0; rowIndex < logicalRows.length; rowIndex += 1) {
    let physicalColumnIndex = 0;
    for (const [logicalColumnIndex, cell] of logicalRows[rowIndex]!.entries()) {
      while (
        physicalColumnIndex < columnCount
        && matrix[rowIndex]![physicalColumnIndex] !== undefined
      ) physicalColumnIndex += 1;
      if (physicalColumnIndex >= columnCount) {
        throw new TypeError('Table rows must cover the same number of physical cells');
      }
      const rowspan = cell.rowspan ?? 1;
      const colspan = cell.colspan ?? 1;
      if (
        rowIndex + rowspan > logicalRows.length
        || physicalColumnIndex + colspan > columnCount
      ) {
        throw new RangeError(
          `Table cell ${rowIndex},${logicalColumnIndex} span is out of range`,
        );
      }
      for (let rowOffset = 0; rowOffset < rowspan; rowOffset += 1) {
        for (let columnOffset = 0; columnOffset < colspan; columnOffset += 1) {
          const targetRow = rowIndex + rowOffset;
          const targetColumn = physicalColumnIndex + columnOffset;
          if (matrix[targetRow]![targetColumn] !== undefined) {
            throw new RangeError(
              `Table cell ${rowIndex},${logicalColumnIndex} span overlaps another cell`,
            );
          }
        }
      }
      for (let rowOffset = 0; rowOffset < rowspan; rowOffset += 1) {
        for (let columnOffset = 0; columnOffset < colspan; columnOffset += 1) {
          const targetRow = rowIndex + rowOffset;
          const targetColumn = physicalColumnIndex + columnOffset;
          matrix[targetRow]![targetColumn] = rowOffset === 0 && columnOffset === 0
            ? cell
            : tableContinuationCell(rowOffset, columnOffset, rowspan, colspan);
        }
      }
      physicalColumnIndex += colspan;
    }
    if (matrix[rowIndex]!.some((cell) => cell === undefined)) {
      throw new TypeError('Table rows must cover the same number of physical cells');
    }
  }
  return matrix as readonly (readonly NormalizedTableCell[])[];
}

function tableContinuationCell(
  rowOffset: number,
  columnOffset: number,
  rowspan: number,
  colspan: number,
): NormalizedTableCell {
  return {
    text: '',
    continuation: {
      ...(rowOffset === 0 && rowspan > 1 ? { rowSpan: rowspan } : {}),
      ...(columnOffset === 0 && colspan > 1 ? { gridSpan: colspan } : {}),
      ...(rowOffset > 0 ? { vertical: true as const } : {}),
      ...(columnOffset > 0 ? { horizontal: true as const } : {}),
    },
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

export function renderEmptyTableCellFragment(
  tokens: Readonly<TableCellMergeTokens> = {},
): string {
  const cell = renderTableCell({ text: '' }, undefined);
  const attributes = [
    renderStructuralMergeSpan(tokens.rowSpan, 'rowSpan'),
    renderStructuralMergeSpan(tokens.gridSpan, 'gridSpan'),
    tokens.vertical ? ' vMerge="1"' : '',
    tokens.horizontal ? ' hMerge="1"' : '',
  ].join('');
  return attributes === ''
    ? cell
    : `<a:tc${attributes}${cell.slice('<a:tc'.length)}`;
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

function readDenseArray(
  value: unknown,
  context: string,
  allowEmpty = false,
): readonly unknown[] {
  if (!Array.isArray(value) || (!allowEmpty && value.length === 0)) {
    throw new TypeError(
      allowEmpty ? `${context} must be an array` : `${context} must be a non-empty array`,
    );
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
  const spanAttributes = renderTableCellSpanAttributes(cell);
  if (cell.continuation !== undefined) {
    if (
      hyperlinkRelationshipId !== undefined
      || runHyperlinkRelationshipIds?.some((paragraph) =>
        paragraph.some((relationshipId) => relationshipId !== undefined))
    ) {
      throw new TypeError('Table continuation cells cannot contain hyperlink relationships');
    }
    return `<a:tc${spanAttributes}><a:tcPr/></a:tc>`;
  }
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
    ...(cell.fontFamily === undefined
      ? {}
      : { defaultFontFamily: cell.fontFamily }),
    ...(cell.fontSize === undefined ? {} : { defaultFontSize: cell.fontSize }),
    ...(cell.bold === undefined ? {} : { defaultBold: cell.bold }),
    ...(cell.color === undefined
      ? {}
      : {
          defaultColor: cell.color,
          suppressDefaultColorForHyperlinks: true,
        }),
    ...(cell.spacing === undefined ? {} : { defaultSpacing: cell.spacing }),
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
  return `<a:tc${spanAttributes}><a:txBody>${bodyProperties}<a:lstStyle/>${renderedParagraphs}</a:txBody><a:tcPr${marginAttributes}${verticalAlignmentAttribute}${textDirectionAttribute}>${borders}${fill}</a:tcPr></a:tc>`;
}

function renderStructuralMergeSpan(
  value: number | undefined,
  name: 'rowSpan' | 'gridSpan',
): string {
  if (value === undefined || value === 1) return '';
  if (!Number.isSafeInteger(value) || value < 1) {
    throw new RangeError(`Table cell ${name} must be a positive safe integer`);
  }
  return ` ${name}="${value}"`;
}

function renderTableCellSpanAttributes(cell: NormalizedTableCell): string {
  const continuation = cell.continuation;
  return [
    continuation?.rowSpan !== undefined
      ? ` rowSpan="${continuation.rowSpan}"`
      : cell.rowspan !== undefined ? ` rowSpan="${cell.rowspan}"` : '',
    continuation?.gridSpan !== undefined
      ? ` gridSpan="${continuation.gridSpan}"`
      : cell.colspan !== undefined ? ` gridSpan="${cell.colspan}"` : '',
    continuation?.vertical ? ' vMerge="1"' : '',
    continuation?.horizontal ? ' hMerge="1"' : '',
  ].join('');
}

function tableCellRichText(
  cell: NormalizedTableCell,
): ReturnType<typeof normalizeRichText> {
  return cell.richText ?? normalizeRichText([
    { runs: [{ text: cell.text, style: {} }] },
  ]);
}

function tableCellHasRenderedHyperlink(cell: NormalizedTableCell): boolean {
  if (cell.continuation !== undefined) return false;
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
