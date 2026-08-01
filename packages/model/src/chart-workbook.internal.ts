import {
  escapeXmlText,
  LosslessXmlDocument,
  type XmlAttribute,
  type XmlElement,
} from '@pptx/lossless-xml';
import { OpcPackage } from '@pptx/opc';
import type { ChartCategories, ChartDefinition, ChartSeries } from './chart.js';

const FIXED_ZIP_DATE = new Date('1980-01-01T00:00:00.000Z');
const SPREADSHEET_NAMESPACE =
  'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OFFICE_DOCUMENT_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const WORKSHEET_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
const STYLES_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
const WORKBOOK_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';
const WORKSHEET_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
const STYLES_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml';
const DECIMAL_PATTERN = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$/;
const CELL_REFERENCE_PATTERN = /^([A-Z]+)([1-9]\d*)$/;

const WORKBOOK_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
  + `<workbook xmlns="${SPREADSHEET_NAMESPACE}" xmlns:r="${RELATIONSHIP_NAMESPACE}">`
  + '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>'
  + '</workbook>';

const STYLES_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
  + `<styleSheet xmlns="${SPREADSHEET_NAMESPACE}">`
  + '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
  + '<fills count="2"><fill><patternFill patternType="none"/></fill>'
  + '<fill><patternFill patternType="gray125"/></fill></fills>'
  + '<borders count="1"><border/></borders>'
  + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
  + '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
  + '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
  + '</styleSheet>';

type WorkbookCell = string | number;

export interface ChartWorkbookFormula {
  readonly groupIndex: number;
  readonly seriesIndex: number;
  readonly name: string;
  readonly categories?: readonly string[];
  readonly values: string;
  readonly xValues?: string;
  readonly sizes?: string;
}

export interface ChartWorkbookPlan {
  readonly worksheetXml: string;
  readonly formulas: readonly Readonly<ChartWorkbookFormula>[];
}

interface InternalWorkbookPlan extends ChartWorkbookPlan {
  readonly cells: readonly (readonly WorkbookCell[])[];
}

interface CategoryBlock {
  readonly key: string;
  readonly ranges: readonly string[];
}

export function planChartWorkbook(
  definition: Readonly<ChartDefinition>,
): Readonly<ChartWorkbookPlan> {
  const plan = createWorkbookPlan(definition);
  return Object.freeze({ worksheetXml: plan.worksheetXml, formulas: plan.formulas });
}

export async function buildChartWorkbook(
  definition: Readonly<ChartDefinition>,
): Promise<Uint8Array> {
  const plan = createWorkbookPlan(definition);
  const workbook = OpcPackage.create({ entryDate: FIXED_ZIP_DATE });
  workbook.transaction(() => {
    workbook.setPart('/xl/workbook.xml', WORKBOOK_XML, WORKBOOK_CONTENT_TYPE);
    workbook.setPart('/xl/worksheets/sheet1.xml', plan.worksheetXml, WORKSHEET_CONTENT_TYPE);
    workbook.setPart('/xl/styles.xml', STYLES_XML, STYLES_CONTENT_TYPE);
    workbook.addRelationship('/', {
      id: 'rId1',
      type: OFFICE_DOCUMENT_RELATIONSHIP,
      target: 'xl/workbook.xml',
    });
    workbook.addRelationship('/xl/workbook.xml', {
      id: 'rId1',
      type: WORKSHEET_RELATIONSHIP,
      target: 'worksheets/sheet1.xml',
    });
    workbook.addRelationship('/xl/workbook.xml', {
      id: 'rId2',
      type: STYLES_RELATIONSHIP,
      target: 'styles.xml',
    });
  });
  return workbook.write();
}

export async function readChartWorkbookCells(
  bytes: Uint8Array,
): Promise<readonly (readonly WorkbookCell[])[]> {
  const workbook = await OpcPackage.open(bytes);
  const sheetUri = readSheetUri(workbook);
  if (workbook.hasPart('/xl/sharedStrings.xml')) {
    throw new Error('Chart workbook shared strings are unsupported');
  }
  if (workbook.parts.some(({ uri }) => uri.startsWith('/xl/externalLinks/'))) {
    throw new Error('Chart workbook external links are unsupported');
  }
  const sheetPart = workbook.requirePart(sheetUri);
  if (sheetPart.contentType !== WORKSHEET_CONTENT_TYPE) {
    throw new Error('Chart workbook worksheet content type is unsupported');
  }
  return readWorksheetCells(LosslessXmlDocument.parse(sheetPart.bytes));
}

export async function chartWorkbookMatches(
  bytes: Uint8Array,
  definition: Readonly<ChartDefinition>,
): Promise<boolean> {
  try {
    const actual = await readChartWorkbookCells(bytes);
    return equalCells(actual, createWorkbookPlan(definition).cells);
  } catch {
    return false;
  }
}

function createWorkbookPlan(definition: Readonly<ChartDefinition>): InternalWorkbookPlan {
  const columns: WorkbookCell[][] = [];
  const formulas: Readonly<ChartWorkbookFormula>[] = [];
  definition.groups.forEach((group, groupIndex) => {
    if (group.type === 'scatter' || group.type === 'bubble') {
      group.series.forEach((series, seriesIndex) => {
        const xColumn = columns.length + 1;
        columns.push(['', ...series.xValues!]);
        const valueColumn = columns.length + 1;
        columns.push([series.name, ...series.values]);
        let sizeColumn: number | undefined;
        if (group.type === 'bubble') {
          sizeColumn = columns.length + 1;
          columns.push(['', ...series.sizes!]);
        }
        formulas.push(Object.freeze({
          groupIndex,
          seriesIndex,
          name: formulaCell(valueColumn, 1),
          values: formulaRange(valueColumn, series.values.length),
          xValues: formulaRange(xColumn, series.xValues!.length),
          ...(sizeColumn === undefined
            ? {}
            : { sizes: formulaRange(sizeColumn, series.sizes!.length) }),
        }));
      });
      return;
    }
    planCategoricalGroup(columns, formulas, group.series, groupIndex);
  });
  const cells = freezeCells(transposeColumns(columns));
  const frozenFormulas = Object.freeze(formulas);
  return Object.freeze({
    worksheetXml: renderWorksheet(cells),
    formulas: frozenFormulas,
    cells,
  });
}

function planCategoricalGroup(
  columns: WorkbookCell[][],
  formulas: Readonly<ChartWorkbookFormula>[],
  series: readonly Readonly<ChartSeries>[],
  groupIndex: number,
): void {
  const blocks: CategoryBlock[] = [];
  const seriesBlockKeys: string[] = [];
  for (const entry of series) {
    const categories = entry.categories!;
    const key = categoryKey(categories);
    seriesBlockKeys.push(key);
    if (blocks.some((block) => block.key === key)) continue;
    const levels = categoryLevels(categories);
    const firstColumn = columns.length + 1;
    levels.forEach((level) => columns.push(['', ...level]));
    blocks.push({
      key,
      ranges: Object.freeze(levels.map((level, index) =>
        formulaRange(firstColumn + index, level.length))),
    });
  }
  series.forEach((entry, seriesIndex) => {
    const valueColumn = columns.length + 1;
    columns.push([entry.name, ...entry.values]);
    const block = blocks.find(({ key }) => key === seriesBlockKeys[seriesIndex])!;
    formulas.push(Object.freeze({
      groupIndex,
      seriesIndex,
      name: formulaCell(valueColumn, 1),
      categories: block.ranges,
      values: formulaRange(valueColumn, entry.values.length),
    }));
  });
}

function categoryLevels(categories: ChartCategories): readonly (readonly WorkbookCell[])[] {
  const first = categories[0];
  return Array.isArray(first)
    ? categories as readonly (readonly string[])[]
    : [categories as readonly WorkbookCell[]];
}

function categoryKey(categories: ChartCategories): string {
  return JSON.stringify(categories);
}

function transposeColumns(columns: readonly (readonly WorkbookCell[])[]): WorkbookCell[][] {
  const rowCount = Math.max(...columns.map((column) => column.length));
  return Array.from({ length: rowCount }, (_row, rowIndex) =>
    columns.map((column) => column[rowIndex] ?? ''));
}

function freezeCells(cells: WorkbookCell[][]): readonly (readonly WorkbookCell[])[] {
  return Object.freeze(cells.map((row) => Object.freeze(row)));
}

function renderWorksheet(cells: readonly (readonly WorkbookCell[])[]): string {
  const rows = cells.map((row, rowIndex) => {
    const rowNumber = rowIndex + 1;
    const renderedCells = row.map((value, columnIndex) =>
      renderCell(columnIndex + 1, rowNumber, value)).join('');
    return `<row r="${rowNumber}">${renderedCells}</row>`;
  }).join('');
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + `<worksheet xmlns="${SPREADSHEET_NAMESPACE}"><sheetData>${rows}</sheetData></worksheet>`;
}

function renderCell(column: number, row: number, value: WorkbookCell): string {
  const reference = `${excelColumn(column)}${row}`;
  if (typeof value === 'number') return `<c r="${reference}"><v>${value}</v></c>`;
  const preserve = /^\s|\s$/.test(value) ? ' xml:space="preserve"' : '';
  return `<c r="${reference}" t="inlineStr"><is><t${preserve}>${escapeXmlText(value)}</t></is></c>`;
}

function formulaCell(column: number, row: number): string {
  return `Sheet1!$${excelColumn(column)}$${row}`;
}

function formulaRange(column: number, count: number): string {
  const name = excelColumn(column);
  return `Sheet1!$${name}$2:$${name}$${count + 1}`;
}

function excelColumn(index: number): string {
  let value = index;
  let result = '';
  while (value > 0) {
    value -= 1;
    result = String.fromCharCode(65 + (value % 26)) + result;
    value = Math.floor(value / 26);
  }
  return result;
}

function readSheetUri(workbook: OpcPackage): string {
  const rootRelationships = workbook.relationships('/');
  const officeRelationships = rootRelationships.filter(
    ({ type }) => type === OFFICE_DOCUMENT_RELATIONSHIP,
  );
  if (
    rootRelationships.length !== 1
    || officeRelationships.length !== 1
    || officeRelationships[0]!.targetMode !== 'Internal'
    || !officeRelationships[0]!.resolvedTarget
  ) {
    throw new Error('Chart workbook office-document relationship is invalid');
  }
  const workbookUri = officeRelationships[0]!.resolvedTarget;
  const workbookPart = workbook.requirePart(workbookUri);
  if (workbookPart.contentType !== WORKBOOK_CONTENT_TYPE) {
    throw new Error('Chart workbook content type is unsupported');
  }
  const xml = LosslessXmlDocument.parse(workbookPart.bytes);
  const root = requireSingleRoot(xml, 'workbook');
  if (elementChildren(root).length !== 1) {
    throw new Error('Chart workbook contains unsupported structures');
  }
  const sheets = oneDirectChild(root, 'sheets');
  if (elementChildren(sheets).length !== 1) {
    throw new Error('Chart workbook sheets collection is unsupported');
  }
  const sheet = oneDirectChild(sheets, 'sheet');
  if (
    unqualifiedAttribute(sheet, 'name')?.value !== 'Sheet1'
    || unqualifiedAttribute(sheet, 'sheetId')?.value !== '1'
  ) {
    throw new Error('Chart workbook must contain one canonical Sheet1');
  }
  const ids = attributesByExpandedName(sheet, RELATIONSHIP_NAMESPACE, 'id');
  if (ids.length !== 1) throw new Error('Chart workbook Sheet1 relationship id is invalid');

  const relationships = workbook.relationships(workbookUri);
  const worksheetRelationships = relationships.filter(({ type }) => type === WORKSHEET_RELATIONSHIP);
  const styleRelationships = relationships.filter(({ type }) => type === STYLES_RELATIONSHIP);
  if (
    relationships.length !== 2
    || worksheetRelationships.length !== 1
    || styleRelationships.length !== 1
    || worksheetRelationships[0]!.id !== ids[0]!.value
    || worksheetRelationships[0]!.targetMode !== 'Internal'
    || !worksheetRelationships[0]!.resolvedTarget
    || styleRelationships[0]!.targetMode !== 'Internal'
    || !styleRelationships[0]!.resolvedTarget
  ) {
    throw new Error('Chart workbook relationships are unsupported');
  }
  const styles = workbook.requirePart(styleRelationships[0]!.resolvedTarget);
  if (styles.contentType !== STYLES_CONTENT_TYPE) {
    throw new Error('Chart workbook styles content type is unsupported');
  }
  return worksheetRelationships[0]!.resolvedTarget;
}

function readWorksheetCells(
  xml: LosslessXmlDocument,
): readonly (readonly WorkbookCell[])[] {
  const root = requireSingleRoot(xml, 'worksheet');
  const rootChildren = elementChildren(root);
  if (rootChildren.length !== 1 || rootChildren[0]!.localName !== 'sheetData') {
    throw new Error('Chart worksheet contains unsupported structures');
  }
  const sheetData = oneDirectChild(root, 'sheetData');
  const rows = elementChildren(sheetData);
  if (
    rows.length === 0
    || rows.some((row) => row.localName !== 'row'
      || elementNamespaceUri(row) !== SPREADSHEET_NAMESPACE)
  ) {
    throw new Error('Chart worksheet rows are invalid');
  }
  const result: WorkbookCell[][] = [];
  let columnCount: number | undefined;
  rows.forEach((row, rowIndex) => {
    const rowNumber = rowIndex + 1;
    const rowReference = unqualifiedAttribute(row, 'r')?.value;
    if (rowReference !== String(rowNumber)) throw new Error('Chart worksheet rows must be contiguous');
    const cells = elementChildren(row);
    if (
      cells.length === 0
      || cells.some((cell) => cell.localName !== 'c'
        || elementNamespaceUri(cell) !== SPREADSHEET_NAMESPACE)
    ) {
      throw new Error('Chart worksheet cells are invalid');
    }
    if (columnCount === undefined) columnCount = cells.length;
    if (cells.length !== columnCount) throw new Error('Chart worksheet has missing cells');
    result.push(cells.map((cell, columnIndex) =>
      readCell(xml, cell, columnIndex + 1, rowNumber)));
  });
  return freezeCells(result);
}

function readCell(
  xml: LosslessXmlDocument,
  cell: XmlElement,
  expectedColumn: number,
  expectedRow: number,
): WorkbookCell {
  const reference = unqualifiedAttribute(cell, 'r')?.value;
  const match = reference ? CELL_REFERENCE_PATTERN.exec(reference) : undefined;
  if (
    !match
    || excelColumnIndex(match[1]!) !== expectedColumn
    || Number(match[2]) !== expectedRow
  ) {
    throw new Error('Chart worksheet cell references must be dense and ordered');
  }
  const children = elementChildren(cell);
  if (children.some(({ localName }) => localName === 'f')) {
    throw new Error('Chart worksheet formulas are unsupported');
  }
  const type = unqualifiedAttribute(cell, 't')?.value;
  if (type === 'inlineStr') {
    if (
      children.length !== 1
      || children[0]!.localName !== 'is'
      || elementNamespaceUri(children[0]!) !== SPREADSHEET_NAMESPACE
    ) {
      throw new Error('Chart worksheet inline string is invalid');
    }
    const text = oneDirectChild(children[0]!, 't');
    if (elementChildren(text).length > 0) throw new Error('Chart worksheet rich strings are unsupported');
    return xml.text(text);
  }
  if (type !== undefined) throw new Error('Chart worksheet cell type is unsupported');
  if (
    children.length !== 1
    || children[0]!.localName !== 'v'
    || elementNamespaceUri(children[0]!) !== SPREADSHEET_NAMESPACE
  ) {
    throw new Error('Chart worksheet numeric cell is invalid');
  }
  const lexical = xml.text(children[0]!);
  if (!DECIMAL_PATTERN.test(lexical)) throw new Error('Chart worksheet number is invalid');
  const value = Number(lexical);
  if (!Number.isFinite(value)) throw new Error('Chart worksheet number is not finite');
  return Object.is(value, -0) ? 0 : value;
}

function requireSingleRoot(xml: LosslessXmlDocument, localName: string): XmlElement {
  if (xml.roots.length !== 1) throw new Error(`Chart workbook ${localName} root is ambiguous`);
  const root = xml.roots[0]!;
  if (root.localName !== localName || elementNamespaceUri(root) !== SPREADSHEET_NAMESPACE) {
    throw new Error(`Chart workbook ${localName} root is invalid`);
  }
  return root;
}

function oneDirectChild(parent: XmlElement, localName: string): XmlElement {
  const children = elementChildren(parent).filter((child) =>
    child.localName === localName && elementNamespaceUri(child) === SPREADSHEET_NAMESPACE);
  if (children.length !== 1) throw new Error(`${localName} must occur exactly once`);
  return children[0]!;
}

function elementChildren(parent: XmlElement): XmlElement[] {
  return parent.children.filter((child): child is XmlElement => child.type === 'element');
}

function unqualifiedAttribute(element: XmlElement, name: string): XmlAttribute | undefined {
  const attributes = element.attributes.filter((attribute) => attribute.name === name);
  if (attributes.length > 1) throw new Error(`Attribute ${name} is ambiguous`);
  return attributes[0];
}

function attributesByExpandedName(
  element: XmlElement,
  namespace: string,
  localName: string,
): readonly XmlAttribute[] {
  return element.attributes.filter((attribute) =>
    attribute.localName === localName && attributeNamespaceUri(element, attribute) === namespace);
}

function elementNamespaceUri(element: XmlElement): string | undefined {
  return namespaceUriForPrefix(element, lexicalPrefix(element.name));
}

function attributeNamespaceUri(element: XmlElement, attribute: XmlAttribute): string | undefined {
  const prefix = lexicalPrefix(attribute.name);
  return prefix === '' ? undefined : namespaceUriForPrefix(element, prefix);
}

function namespaceUriForPrefix(element: XmlElement, prefix: string): string | undefined {
  for (let scope: XmlElement | undefined = element; scope; scope = scope.parent) {
    for (const attribute of scope.attributes) {
      if (attribute.name === 'xmlns' && prefix === '') return attribute.value;
      if (attribute.name === `xmlns:${prefix}`) return attribute.value;
    }
  }
  return undefined;
}

function lexicalPrefix(name: string): string {
  const separator = name.indexOf(':');
  return separator < 0 ? '' : name.slice(0, separator);
}

function excelColumnIndex(value: string): number {
  let result = 0;
  for (const character of value) result = result * 26 + character.charCodeAt(0) - 64;
  return result;
}

function equalCells(
  left: readonly (readonly WorkbookCell[])[],
  right: readonly (readonly WorkbookCell[])[],
): boolean {
  return left.length === right.length && left.every((row, rowIndex) => {
    const other = right[rowIndex];
    return other !== undefined
      && row.length === other.length
      && row.every((cell, columnIndex) => cell === other[columnIndex]);
  });
}
