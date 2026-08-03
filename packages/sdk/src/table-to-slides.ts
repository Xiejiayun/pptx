import type {
  AddShapeOptions,
  AddTableCellInput,
  AddTableCellOptions,
  AddTableOptions,
  AddTextOptions,
  PresetShapeType,
  RichTextParagraph,
  TableAutoPageMarginInput,
} from '@pptx/model';
import type { AddImageSourceOptions, ImageSource } from './raster-image-source.js';
import { mapComputedCellOptions } from './table-to-slides-css.js';

const OPTION_KEYS = new Set([
  'name',
  'masterSlideName',
  'autoPage',
  'autoPageCharWeight',
  'autoPageLineWeight',
  'autoPageRepeatHeader',
  'autoPageHeaderRows',
  'autoPageSlideStartY',
  'slideMargin',
  'x',
  'y',
  'width',
  'height',
  'columnWidths',
  'addImage',
  'addShape',
  'addTable',
  'addText',
]);
const AUTO_PAGE_CONTROL_KEYS = [
  'autoPageCharWeight',
  'autoPageLineWeight',
  'autoPageRepeatHeader',
  'autoPageHeaderRows',
  'autoPageSlideStartY',
  'slideMargin',
] as const;
export interface TableToSlidesAddImage {
  readonly source: ImageSource;
  readonly options?: AddImageSourceOptions;
}

export interface TableToSlidesAddShape {
  readonly type: PresetShapeType;
  readonly options?: AddShapeOptions;
}

export interface TableToSlidesAddTable {
  readonly rows: readonly (readonly AddTableCellInput[])[];
  readonly options?: AddTableOptions;
}

export interface TableToSlidesAddText {
  readonly text: string | readonly RichTextParagraph[];
  readonly options?: AddTextOptions;
}

export interface TableToSlidesOptions {
  readonly name?: string;
  readonly masterSlideName?: string;
  readonly autoPage?: boolean;
  readonly autoPageCharWeight?: number;
  readonly autoPageLineWeight?: number;
  readonly autoPageRepeatHeader?: boolean;
  readonly autoPageHeaderRows?: number;
  readonly autoPageSlideStartY?: number;
  readonly slideMargin?: TableAutoPageMarginInput;
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
  readonly columnWidths?: number | readonly number[];
  readonly addImage?: TableToSlidesAddImage;
  readonly addShape?: TableToSlidesAddShape;
  readonly addTable?: TableToSlidesAddTable;
  readonly addText?: TableToSlidesAddText;
}

export interface NormalizedTableToSlidesRequest extends TableToSlidesOptions {
  readonly id: string;
  readonly autoPage: boolean;
}

export interface HtmlTableCellSnapshot {
  readonly text: string;
  readonly offsetWidth: number;
  readonly header: boolean;
  readonly colspan?: number;
  readonly rowspan?: number;
  readonly pptxWidth?: string;
  readonly pptxMinWidth?: string;
  readonly options: Readonly<AddTableCellOptions>;
}

export interface HtmlTableSnapshot {
  readonly rows: readonly (readonly HtmlTableCellSnapshot[])[];
  readonly headRowCount: number;
  readonly widthSourceRowIndex: number;
}

export type HtmlComputedStyleResolver = (element: unknown) => unknown;

export function normalizeTableToSlidesRequest(
  elementId: unknown,
  options: unknown = undefined,
): Readonly<NormalizedTableToSlidesRequest> {
  const id = normalizeElementId(elementId);
  const input = options === undefined
    ? Object.create(null) as Record<string, unknown>
    : readDataObject(options, OPTION_KEYS, 'HTML table slide options');
  const autoPage = normalizeOptionalBoolean(input.autoPage, 'HTML table autoPage') ?? true;
  if (!autoPage && AUTO_PAGE_CONTROL_KEYS.some((key) => input[key] !== undefined)) {
    throw new TypeError('HTML table auto-page controls require autoPage to be true');
  }

  const name = normalizeOptionalString(input.name, 'HTML table name');
  const masterSlideName = normalizeOptionalNonEmptyString(
    input.masterSlideName,
    'HTML table masterSlideName',
  );
  const autoPageCharWeight = normalizeOptionalWeight(
    input.autoPageCharWeight,
    'HTML table autoPageCharWeight',
  );
  const autoPageLineWeight = normalizeOptionalWeight(
    input.autoPageLineWeight,
    'HTML table autoPageLineWeight',
  );
  const autoPageRepeatHeader = normalizeOptionalBoolean(
    input.autoPageRepeatHeader,
    'HTML table autoPageRepeatHeader',
  );
  const autoPageHeaderRows = normalizeOptionalPositiveSafeInteger(
    input.autoPageHeaderRows,
    'HTML table autoPageHeaderRows',
  );
  if (autoPageHeaderRows !== undefined && autoPageRepeatHeader !== true) {
    throw new TypeError(
      'HTML table autoPageHeaderRows requires autoPageRepeatHeader to be true',
    );
  }
  const autoPageSlideStartY = normalizeOptionalNonNegativeSafeInteger(
    input.autoPageSlideStartY,
    'HTML table autoPageSlideStartY',
  );
  const slideMargin = input.slideMargin === undefined
    ? undefined
    : normalizeSlideMargin(input.slideMargin);
  const x = normalizeOptionalCoordinate(input.x, 'HTML table x');
  const y = normalizeOptionalCoordinate(input.y, 'HTML table y');
  const width = normalizeOptionalPositiveCoordinate(input.width, 'HTML table width');
  const height = normalizeOptionalPositiveCoordinate(input.height, 'HTML table height');
  const columnWidths = input.columnWidths === undefined
    ? undefined
    : normalizeColumnWidths(input.columnWidths);
  const addImage = input.addImage === undefined
    ? undefined
    : normalizeAddImage(input.addImage);
  const addShape = input.addShape === undefined
    ? undefined
    : normalizeAddShape(input.addShape);
  const addTable = input.addTable === undefined
    ? undefined
    : normalizeAddTable(input.addTable);
  const addText = input.addText === undefined
    ? undefined
    : normalizeAddText(input.addText);

  return Object.freeze({
    id,
    autoPage,
    ...(name === undefined ? {} : { name }),
    ...(masterSlideName === undefined ? {} : { masterSlideName }),
    ...(autoPageCharWeight === undefined ? {} : { autoPageCharWeight }),
    ...(autoPageLineWeight === undefined ? {} : { autoPageLineWeight }),
    ...(autoPageRepeatHeader === undefined ? {} : { autoPageRepeatHeader }),
    ...(autoPageHeaderRows === undefined ? {} : { autoPageHeaderRows }),
    ...(autoPageSlideStartY === undefined ? {} : { autoPageSlideStartY }),
    ...(slideMargin === undefined ? {} : { slideMargin }),
    ...(x === undefined ? {} : { x }),
    ...(y === undefined ? {} : { y }),
    ...(width === undefined ? {} : { width }),
    ...(height === undefined ? {} : { height }),
    ...(columnWidths === undefined ? {} : { columnWidths }),
    ...(addImage === undefined ? {} : { addImage }),
    ...(addShape === undefined ? {} : { addShape }),
    ...(addTable === undefined ? {} : { addTable }),
    ...(addText === undefined ? {} : { addText }),
  });
}

export function snapshotHtmlTableById(
  elementId: unknown,
  documentValue: unknown = typeof document === 'undefined' ? undefined : document,
): Readonly<HtmlTableSnapshot> {
  const id = normalizeElementId(elementId);
  if (!documentValue || typeof documentValue !== 'object') {
    throw new TypeError('HTML table import requires a browser document');
  }
  const getElementById = readMethod(documentValue, 'getElementById');
  if (!getElementById) {
    throw new TypeError('HTML table browser document must provide getElementById()');
  }
  const target = callPlatformMethod(
    getElementById,
    documentValue,
    [id],
    'HTML table document getElementById',
  );
  if (target === null || target === undefined) {
    throw new RangeError(`HTML table element ${JSON.stringify(id)} was not found`);
  }
  requireTable(target, `HTML table element ${JSON.stringify(id)}`);
  const ownerDocument = readProperty(target, 'ownerDocument', 'HTML table ownerDocument');
  if (!ownerDocument || typeof ownerDocument !== 'object') {
    throw new TypeError('HTML table ownerDocument must be available');
  }
  const defaultView = readProperty(ownerDocument, 'defaultView', 'HTML table defaultView');
  if (!defaultView || typeof defaultView !== 'object') {
    throw new TypeError('HTML table ownerDocument defaultView must be available');
  }
  const getComputedStyle = readMethod(defaultView, 'getComputedStyle');
  if (!getComputedStyle) {
    throw new TypeError('HTML table defaultView must provide getComputedStyle()');
  }
  return snapshotHtmlTable(
    target,
    (element) => callPlatformMethod(
      getComputedStyle,
      defaultView,
      [element],
      'HTML table defaultView getComputedStyle',
    ),
  );
}

export function snapshotHtmlTable(
  table: unknown,
  getComputedStyle: HtmlComputedStyleResolver,
): Readonly<HtmlTableSnapshot> {
  requireTable(table, 'HTML table');
  if (typeof getComputedStyle !== 'function') {
    throw new TypeError('HTML table getComputedStyle must be a function');
  }

  const headRows = readSectionRows(
    readProperty(table, 'tHead', 'HTML table tHead'),
    'HTML table thead',
  );
  const bodySections = readArrayLike(
    readProperty(table, 'tBodies', 'HTML table tBodies'),
    'HTML table tbody collection',
  );
  const bodyRows = bodySections.flatMap((section, index) =>
    readSectionRows(section, `HTML table tbody ${index}`));
  const footRows = readSectionRows(
    readProperty(table, 'tFoot', 'HTML table tFoot'),
    'HTML table tfoot',
  );
  const platformRows = [...headRows, ...bodyRows, ...footRows];
  if (platformRows.length === 0) {
    throw new RangeError('HTML table must contain at least one row');
  }

  const rows = platformRows.map((row, rowIndex) => {
    if (!row || typeof row !== 'object') {
      throw new TypeError(`HTML table row ${rowIndex} must be an object`);
    }
    const cells = readArrayLike(
      readProperty(row, 'cells', `HTML table row ${rowIndex} cells`),
      `HTML table row ${rowIndex} cell collection`,
    );
    if (cells.length === 0) {
      throw new RangeError(`HTML table row ${rowIndex} must contain at least one cell`);
    }
    return Object.freeze(cells.map((cell, cellIndex) =>
      snapshotCell(cell, rowIndex, cellIndex, getComputedStyle)));
  });

  return Object.freeze({
    rows: Object.freeze(rows),
    headRowCount: headRows.length,
    widthSourceRowIndex: 0,
  });
}

function snapshotCell(
  cell: unknown,
  rowIndex: number,
  cellIndex: number,
  getComputedStyle: HtmlComputedStyleResolver,
): Readonly<HtmlTableCellSnapshot> {
  const context = `HTML table cell ${rowIndex}:${cellIndex}`;
  if (!cell || typeof cell !== 'object') {
    throw new TypeError(`${context} must be an object`);
  }
  const localName = readProperty(cell, 'localName', `${context} localName`);
  if (localName !== 'td' && localName !== 'th') {
    throw new TypeError(`${context} localName must be td or th`);
  }
  const text = readProperty(cell, 'innerText', `${context} innerText`);
  if (typeof text !== 'string') throw new TypeError(`${context} innerText must be a string`);
  const offsetWidth = readProperty(cell, 'offsetWidth', `${context} offsetWidth`);
  if (typeof offsetWidth !== 'number' || !Number.isFinite(offsetWidth) || offsetWidth < 0) {
    throw new TypeError(`${context} offsetWidth must be a non-negative finite number`);
  }
  const colspan = normalizePlatformSpan(
    readProperty(cell, 'colSpan', `${context} colSpan`),
    `${context} colSpan`,
  );
  const rowspan = normalizePlatformSpan(
    readProperty(cell, 'rowSpan', `${context} rowSpan`),
    `${context} rowSpan`,
  );
  const getAttribute = readMethod(cell, 'getAttribute');
  if (!getAttribute) throw new TypeError(`${context} must provide getAttribute()`);
  const pptxWidth = normalizePlatformAttribute(
    callPlatformMethod(
      getAttribute,
      cell,
      ['data-pptx-width'],
      `${context} getAttribute(data-pptx-width)`,
    ),
    `${context} data-pptx-width`,
  );
  const pptxMinWidth = normalizePlatformAttribute(
    callPlatformMethod(
      getAttribute,
      cell,
      ['data-pptx-min-width'],
      `${context} getAttribute(data-pptx-min-width)`,
    ),
    `${context} data-pptx-min-width`,
  );
  let style: unknown;
  try {
    style = getComputedStyle(cell);
  } catch (error) {
    throw new TypeError(`${context} getComputedStyle failed`, { cause: error });
  }
  if (!style || (typeof style !== 'object' && typeof style !== 'function')) {
    throw new TypeError(`${context} getComputedStyle must return an object`);
  }
  const options = mapComputedCellOptions(style, context);

  return Object.freeze({
    text: text.replace(/\r\n?/g, '\n'),
    offsetWidth,
    header: localName === 'th',
    ...(colspan === undefined ? {} : { colspan }),
    ...(rowspan === undefined ? {} : { rowspan }),
    ...(pptxWidth === undefined ? {} : { pptxWidth }),
    ...(pptxMinWidth === undefined ? {} : { pptxMinWidth }),
    options,
  });
}

function normalizeElementId(value: unknown): string {
  if (typeof value !== 'string' || value.length === 0) {
    throw new TypeError('HTML table element ID must be a non-empty string');
  }
  return value;
}

function normalizeOptionalString(value: unknown, context: string): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'string') throw new TypeError(`${context} must be a string`);
  return value;
}

function normalizeOptionalNonEmptyString(
  value: unknown,
  context: string,
): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'string' || value.length === 0) {
    throw new TypeError(`${context} must be a non-empty string`);
  }
  return value;
}

function normalizeOptionalBoolean(value: unknown, context: string): boolean | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'boolean') throw new TypeError(`${context} must be a boolean`);
  return value;
}

function normalizeOptionalWeight(value: unknown, context: string): number | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  if (value < -1 || value > 1) throw new RangeError(`${context} must be between -1 and 1`);
  return Object.is(value, -0) ? 0 : value;
}

function normalizeOptionalCoordinate(value: unknown, context: string): number | undefined {
  if (value === undefined) return undefined;
  return normalizeCoordinate(value, context);
}

function normalizeOptionalPositiveCoordinate(
  value: unknown,
  context: string,
): number | undefined {
  if (value === undefined) return undefined;
  const coordinate = normalizeCoordinate(value, context);
  if (coordinate <= 0) throw new RangeError(`${context} must be positive`);
  return coordinate;
}

function normalizeCoordinate(value: unknown, context: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  const rounded = Math.round(value);
  if (!Number.isSafeInteger(rounded)) {
    throw new RangeError(`${context} must round to a safe integer EMU value`);
  }
  return Object.is(rounded, -0) ? 0 : rounded;
}

function normalizeOptionalNonNegativeSafeInteger(
  value: unknown,
  context: string,
): number | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  if (!Number.isSafeInteger(value)) {
    throw new RangeError(`${context} must be a safe integer`);
  }
  if (value < 0) throw new RangeError(`${context} must be non-negative`);
  return value;
}

function normalizeOptionalPositiveSafeInteger(
  value: unknown,
  context: string,
): number | undefined {
  const normalized = normalizeOptionalNonNegativeSafeInteger(value, context);
  if (normalized !== undefined && normalized === 0) {
    throw new RangeError(`${context} must be a positive safe integer`);
  }
  return normalized;
}

function normalizeSlideMargin(value: unknown): TableAutoPageMarginInput {
  if (typeof value === 'number') {
    return requireNonNegativeSafeInteger(value, 'HTML table slideMargin');
  }
  const values = readDensePlainArray(value, 'HTML table slideMargin', 4);
  return Object.freeze(values.map((entry, index) => requireNonNegativeSafeInteger(
    entry,
    `HTML table slideMargin ${['top', 'right', 'bottom', 'left'][index]}`,
  ))) as readonly [number, number, number, number];
}

function normalizeColumnWidths(value: unknown): number | readonly number[] {
  if (Array.isArray(value)) {
    const values = readDensePlainArray(value, 'HTML table columnWidths');
    if (values.length === 0) {
      throw new RangeError('HTML table columnWidths must be a non-empty array');
    }
    return Object.freeze(values.map((entry, index) => {
      const width = normalizeCoordinate(entry, `HTML table columnWidths ${index}`);
      if (width <= 0) {
        throw new RangeError(`HTML table columnWidths ${index} must be positive`);
      }
      return width;
    }));
  }
  const width = normalizeCoordinate(value, 'HTML table columnWidths');
  if (width <= 0) throw new RangeError('HTML table columnWidths must be positive');
  return width;
}

function normalizeAddImage(value: unknown): Readonly<TableToSlidesAddImage> {
  const input = readDataObject(
    value,
    new Set(['source', 'options']),
    'HTML table addImage',
  );
  if (!Object.hasOwn(input, 'source') || input.source === undefined) {
    throw new TypeError('HTML table addImage source is required');
  }
  const options = cloneOptionalOptions(input.options, 'HTML table addImage options');
  return Object.freeze({
    source: input.source as ImageSource,
    ...(options === undefined ? {} : { options: options as AddImageSourceOptions }),
  });
}

function normalizeAddShape(value: unknown): Readonly<TableToSlidesAddShape> {
  const input = readDataObject(
    value,
    new Set(['type', 'options']),
    'HTML table addShape',
  );
  if (typeof input.type !== 'string' || input.type.length === 0) {
    throw new TypeError('HTML table addShape type must be a non-empty string');
  }
  const options = cloneOptionalOptions(input.options, 'HTML table addShape options');
  return Object.freeze({
    type: input.type as PresetShapeType,
    ...(options === undefined ? {} : { options: options as AddShapeOptions }),
  });
}

function normalizeAddTable(value: unknown): Readonly<TableToSlidesAddTable> {
  const input = readDataObject(
    value,
    new Set(['rows', 'options']),
    'HTML table addTable',
  );
  const rows = readDensePlainArray(input.rows, 'HTML table addTable rows');
  if (rows.length === 0) throw new RangeError('HTML table addTable rows cannot be empty');
  const normalizedRows = rows.map((row, rowIndex) => {
    const cells = readDensePlainArray(row, `HTML table addTable row ${rowIndex}`);
    if (cells.length === 0) {
      throw new RangeError(`HTML table addTable row ${rowIndex} cannot be empty`);
    }
    return Object.freeze([...cells]) as readonly AddTableCellInput[];
  });
  const options = cloneOptionalOptions(input.options, 'HTML table addTable options');
  return Object.freeze({
    rows: Object.freeze(normalizedRows),
    ...(options === undefined ? {} : { options: options as AddTableOptions }),
  });
}

function normalizeAddText(value: unknown): Readonly<TableToSlidesAddText> {
  const input = readDataObject(
    value,
    new Set(['text', 'options']),
    'HTML table addText',
  );
  let text: string | readonly RichTextParagraph[];
  if (typeof input.text === 'string') text = input.text;
  else {
    const paragraphs = readDensePlainArray(input.text, 'HTML table addText text');
    if (paragraphs.length === 0) throw new RangeError('HTML table addText text cannot be empty');
    text = Object.freeze([...paragraphs]) as readonly RichTextParagraph[];
  }
  const options = cloneOptionalOptions(input.options, 'HTML table addText options');
  return Object.freeze({
    text,
    ...(options === undefined ? {} : { options: options as AddTextOptions }),
  });
}

function cloneOptionalOptions(
  value: unknown,
  context: string,
): Readonly<Record<string, unknown>> | undefined {
  if (value === undefined) return undefined;
  return Object.freeze(readDataObject(value, undefined, context));
}

function readDataObject(
  value: unknown,
  allowed: ReadonlySet<string> | undefined,
  context: string,
): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${context} must be an object`);
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError(`${context} must be an ordinary object`);
  }
  const result = Object.create(null) as Record<string, unknown>;
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || (allowed !== undefined && !allowed.has(key))) {
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

function readDensePlainArray(
  value: unknown,
  context: string,
  expectedLength?: number,
): readonly unknown[] {
  if (!Array.isArray(value)) throw new TypeError(`${context} must be an array`);
  if (Object.getPrototypeOf(value) !== Array.prototype) {
    throw new TypeError(`${context} must be a plain array`);
  }
  if (expectedLength !== undefined && value.length !== expectedLength) {
    throw new RangeError(`${context} must contain exactly ${expectedLength} values`);
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

function requireNonNegativeSafeInteger(value: unknown, context: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  if (!Number.isSafeInteger(value)) throw new RangeError(`${context} must be a safe integer`);
  if (value < 0) throw new RangeError(`${context} must be non-negative`);
  return value;
}

function normalizePlatformSpan(value: unknown, context: string): number | undefined {
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || value < 1) {
    throw new TypeError(`${context} must be a positive safe integer`);
  }
  return value === 1 ? undefined : value;
}

function normalizePlatformAttribute(value: unknown, context: string): string | undefined {
  if (value === null) return undefined;
  if (typeof value !== 'string') throw new TypeError(`${context} must be a string or null`);
  return value;
}

function requireTable(value: unknown, context: string): asserts value is Record<string, unknown> {
  if (!value || typeof value !== 'object') {
    throw new TypeError(`${context} must be a table element`);
  }
  const localName = readProperty(value, 'localName', `${context} localName`);
  if (localName !== 'table') throw new TypeError(`${context} must be a table element`);
}

function readSectionRows(value: unknown, context: string): readonly unknown[] {
  if (value === undefined || value === null) return [];
  if (typeof value !== 'object') throw new TypeError(`${context} must be an object`);
  return readArrayLike(readProperty(value, 'rows', `${context} rows`), `${context} rows`);
}

function readArrayLike(value: unknown, context: string): unknown[] {
  if (!value || (typeof value !== 'object' && typeof value !== 'function')) {
    throw new TypeError(`${context} must be an array-like collection`);
  }
  const length = readProperty(value, 'length', `${context} length`);
  if (typeof length !== 'number' || !Number.isSafeInteger(length) || length < 0) {
    throw new TypeError(`${context} length must be a non-negative safe integer`);
  }
  return Array.from({ length }, (_, index) =>
    readProperty(value, String(index), `${context} item ${index}`));
}

function readMethod(value: object, key: string): ((...args: unknown[]) => unknown) | undefined {
  const method = readProperty(value, key, key);
  return typeof method === 'function' ? method as (...args: unknown[]) => unknown : undefined;
}

function callPlatformMethod(
  method: (...args: unknown[]) => unknown,
  receiver: object,
  args: readonly unknown[],
  context: string,
): unknown {
  try {
    return Reflect.apply(method, receiver, args);
  } catch (error) {
    throw new TypeError(`${context} failed`, { cause: error });
  }
}

function readProperty(value: object, key: string, context: string): unknown {
  try {
    return Reflect.get(value, key);
  } catch (error) {
    throw new TypeError(`${context} could not be read`, { cause: error });
  }
}
