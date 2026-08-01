import { LosslessXmlDocument, type XmlAttribute, type XmlElement } from '@pptx/lossless-xml';
import type { OpcPackage } from '@pptx/opc';
import type {
  ChartCategories,
  ChartDefinitionInput,
  ChartGroupInput,
  ChartSeriesInput,
  ChartState,
  ChartType,
} from './chart.js';
import { normalizeChartDefinition } from './chart-definition.internal.js';

const CHART_NAMESPACE = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const MODERN_CHART_NAMESPACE = 'http://schemas.microsoft.com/office/drawing/2014/chartex';
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package';
const CHART_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const XLSX_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const GROUP_TYPES: Readonly<Record<string, ChartType>> = Object.freeze({
  areaChart: 'area',
  barChart: 'bar',
  bar3DChart: 'bar3D',
  bubbleChart: 'bubble',
  doughnutChart: 'doughnut',
  lineChart: 'line',
  pieChart: 'pie',
  radarChart: 'radar',
  scatterChart: 'scatter',
});
const AXIS_ELEMENTS = new Set(['catAx', 'dateAx', 'serAx', 'valAx']);
const FORMULA_PATTERN = /^(?:Sheet1|'Sheet1')!\$[A-Z]{1,3}\$[1-9]\d*(?::\$[A-Z]{1,3}\$[1-9]\d*)?$/;
const DECIMAL_PATTERN = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$/;
const CANONICAL_INDEX_PATTERN = /^(?:0|[1-9]\d*)$/;

interface ParsedGroup {
  readonly type: ChartType;
  readonly series: readonly ChartSeriesInput[];
  readonly axisIds: readonly number[];
}

interface ParsedAxis {
  readonly id: number;
  readonly crossId: number;
}

class ChartStateError extends Error {
  constructor(
    readonly status: 'unsupported' | 'ambiguous',
    message: string,
  ) {
    super(message);
  }
}

export function readChartState(pkg: OpcPackage, chartPartUri: string): Readonly<ChartState> {
  const part = pkg.getPart(chartPartUri);
  if (!part) return stateError('unsupported', 'Chart part is missing');
  let xml: LosslessXmlDocument;
  try {
    xml = LosslessXmlDocument.parse(part.bytes);
  } catch {
    return stateError('unsupported', 'Chart part is not valid XML');
  }
  if (
    xml.roots.some((root) => elementNamespaceUri(root) === MODERN_CHART_NAMESPACE)
  ) {
    return Object.freeze({
      status: 'modern',
      reason: 'Chart uses the Office 2016 modern chart namespace',
    });
  }
  if (part.contentType !== CHART_CONTENT_TYPE) {
    return stateError('unsupported', 'Chart part has an unsupported content type');
  }

  try {
    return readStandardChartState(pkg, chartPartUri, xml);
  } catch (error) {
    if (error instanceof ChartStateError) return stateError(error.status, error.message);
    return stateError(
      'unsupported',
      error instanceof Error ? error.message : 'Chart structure is unsupported',
    );
  }
}

function readStandardChartState(
  pkg: OpcPackage,
  chartPartUri: string,
  xml: LosslessXmlDocument,
): Readonly<ChartState> {
  if (xml.roots.length !== 1) ambiguous('Chart part must have exactly one root');
  const root = xml.roots[0]!;
  if (root.localName !== 'chartSpace' || elementNamespaceUri(root) !== CHART_NAMESPACE) {
    unsupported('Chart root is not a standard chartSpace element');
  }
  const chart = oneChartChild(root, 'chart', 'chartSpace must contain exactly one chart');
  const plotArea = oneChartChild(chart, 'plotArea', 'chart must contain exactly one plotArea');
  const directPlotChildren = elementChildren(plotArea);
  const groupElements = directPlotChildren.filter((child) =>
    elementNamespaceUri(child) === CHART_NAMESPACE && GROUP_TYPES[child.localName] !== undefined);
  const unknownGroups = directPlotChildren.filter((child) =>
    elementNamespaceUri(child) === CHART_NAMESPACE
      && child.localName.endsWith('Chart')
      && GROUP_TYPES[child.localName] === undefined);
  if (unknownGroups.length > 0) unsupported(`Unsupported chart group ${unknownGroups[0]!.localName}`);
  if (groupElements.length === 0) unsupported('Chart plotArea has no supported chart groups');

  const parsedGroups = groupElements.map((group) => readGroup(xml, group));
  const groups = assignAxes(xml, plotArea, parsedGroups);
  let definition;
  try {
    definition = normalizeChartDefinition({ groups });
  } catch (error) {
    unsupported(error instanceof Error ? error.message : 'Chart definition is unsupported');
  }

  const externalData = chartChildren(root, 'externalData');
  if (externalData.length > 1) ambiguous('Chart has multiple externalData elements');
  if (externalData.length === 0) {
    return Object.freeze({ status: 'cache-only', definition });
  }
  const workbookPartUri = readWorkbookPartUri(pkg, chartPartUri, externalData[0]!);
  return Object.freeze({ status: 'recognized', definition, workbookPartUri });
}

function readGroup(xml: LosslessXmlDocument, group: XmlElement): ParsedGroup {
  const type = GROUP_TYPES[group.localName]!;
  const seriesElements = chartChildren(group, 'ser');
  if (seriesElements.length === 0) unsupported(`${group.localName} has no series`);
  const indexed = seriesElements.map((series) => ({
    element: series,
    index: readRequiredIndex(xml, series, 'idx'),
    order: readRequiredIndex(xml, series, 'order'),
  }));
  requireUnique(indexed.map(({ index }) => index), `${group.localName} series indexes`);
  requireUnique(indexed.map(({ order }) => order), `${group.localName} series order values`);
  const series = indexed
    .sort((left, right) => left.order - right.order)
    .map(({ element }) => readSeries(xml, element, type));
  const axisIds = chartChildren(group, 'axId').map((element) =>
    readUnsignedVal(xml, element, `${group.localName} axis reference`, true));
  requireUnique(axisIds, `${group.localName} axis references`);
  const expectedAxes = type === 'pie' || type === 'doughnut'
    ? [0]
    : type === 'bar3D'
      ? [2, 3]
      : [2];
  if (!expectedAxes.includes(axisIds.length)) {
    unsupported(`${group.localName} has an unsupported axis reference count`);
  }
  return { type, series, axisIds: Object.freeze(axisIds) };
}

function readSeries(
  xml: LosslessXmlDocument,
  series: XmlElement,
  type: ChartType,
): ChartSeriesInput {
  const nameValues = readStringReference(xml, series, 'tx');
  if (nameValues.length !== 1 || nameValues[0]!.length === 0) {
    unsupported('Chart series name cache must contain one non-empty point');
  }
  const name = nameValues[0]!;
  if (type === 'scatter' || type === 'bubble') {
    const xValues = readNumericReference(xml, series, 'xVal');
    const values = readNumericReference(xml, series, 'yVal');
    const sizes = type === 'bubble'
      ? readNumericReference(xml, series, 'bubbleSize')
      : undefined;
    return {
      name,
      values,
      xValues,
      ...(sizes === undefined ? {} : { sizes }),
    };
  }
  return {
    name,
    categories: readCategories(xml, series),
    values: readNumericReference(xml, series, 'val'),
  };
}

function readCategories(xml: LosslessXmlDocument, series: XmlElement): ChartCategories {
  const category = oneChartChild(series, 'cat', 'Chart series must contain exactly one cat element');
  const references = elementChildren(category).filter((child) =>
    elementNamespaceUri(child) === CHART_NAMESPACE
      && (child.localName === 'strRef'
        || child.localName === 'numRef'
        || child.localName === 'multiLvlStrRef'));
  if (references.length > 1) ambiguous('Chart category has multiple cache references');
  if (references.length === 0) unsupported('Chart category has no supported cache reference');
  const reference = references[0]!;
  validateFormula(xml, reference);
  if (reference.localName === 'strRef') {
    return readCache(xml, oneChartChild(
      reference,
      'strCache',
      'String category reference must contain exactly one strCache',
    ), false) as readonly string[];
  }
  if (reference.localName === 'numRef') {
    return readCache(xml, oneChartChild(
      reference,
      'numCache',
      'Numeric category reference must contain exactly one numCache',
    ), true) as readonly number[];
  }
  const cache = oneChartChild(
    reference,
    'multiLvlStrCache',
    'Multi-level category reference must contain exactly one multiLvlStrCache',
  );
  const count = readPointCount(xml, cache);
  const levels = chartChildren(cache, 'lvl');
  if (levels.length === 0) unsupported('Multi-level category cache must contain levels');
  return Object.freeze(levels.map((level) =>
    Object.freeze(readPoints(xml, level, count, false) as string[])));
}

function readStringReference(
  xml: LosslessXmlDocument,
  parent: XmlElement,
  containerName: string,
): readonly string[] {
  const container = oneChartChild(
    parent,
    containerName,
    `Chart series must contain exactly one ${containerName} element`,
  );
  const reference = oneChartChild(
    container,
    'strRef',
    `${containerName} must contain exactly one strRef`,
  );
  validateFormula(xml, reference);
  const cache = oneChartChild(
    reference,
    'strCache',
    `${containerName} must contain exactly one strCache`,
  );
  return readCache(xml, cache, false) as readonly string[];
}

function readNumericReference(
  xml: LosslessXmlDocument,
  parent: XmlElement,
  containerName: string,
): readonly number[] {
  const container = oneChartChild(
    parent,
    containerName,
    `Chart series must contain exactly one ${containerName} element`,
  );
  const reference = oneChartChild(
    container,
    'numRef',
    `${containerName} must contain exactly one numRef`,
  );
  validateFormula(xml, reference);
  const cache = oneChartChild(
    reference,
    'numCache',
    `${containerName} must contain exactly one numCache`,
  );
  return readCache(xml, cache, true) as readonly number[];
}

function readCache(
  xml: LosslessXmlDocument,
  cache: XmlElement,
  numeric: boolean,
): readonly (string | number)[] {
  const count = readPointCount(xml, cache);
  return Object.freeze(readPoints(xml, cache, count, numeric));
}

function readPointCount(xml: LosslessXmlDocument, cache: XmlElement): number {
  const pointCount = oneChartChild(
    cache,
    'ptCount',
    `${cache.localName} must contain exactly one ptCount`,
  );
  return readUnsignedVal(xml, pointCount, `${cache.localName} point count`, false);
}

function readPoints(
  xml: LosslessXmlDocument,
  parent: XmlElement,
  count: number,
  numeric: boolean,
): (string | number)[] {
  const points = chartChildren(parent, 'pt').map((point) => {
    const index = readUnsignedAttribute(xml, point, 'idx', 'Chart cache point index', false);
    const valueElement = oneChartChild(point, 'v', 'Chart cache point must contain exactly one value');
    if (elementChildren(valueElement).length > 0) unsupported('Chart cache value must contain simple text');
    const text = xml.text(valueElement);
    const value = numeric ? readDecimal(text, 'Chart cache numeric value') : text;
    return { index, value };
  });
  if (points.length !== count) unsupported('Chart cache point count does not match its points');
  requireUnique(points.map(({ index }) => index), 'Chart cache point indexes');
  if (points.some(({ index }) => index >= count)) {
    unsupported('Chart cache point index is outside the declared range');
  }
  return points.sort((left, right) => left.index - right.index).map(({ value }) => value);
}

function validateFormula(xml: LosslessXmlDocument, reference: XmlElement): void {
  const formulas = chartChildren(reference, 'f');
  if (formulas.length > 1) ambiguous('Chart cache reference has multiple formulas');
  if (formulas.length === 0) return;
  const formula = formulas[0]!;
  if (elementChildren(formula).length > 0 || !FORMULA_PATTERN.test(xml.text(formula))) {
    unsupported('Chart cache formula is unsupported');
  }
}

function assignAxes(
  xml: LosslessXmlDocument,
  plotArea: XmlElement,
  groups: readonly ParsedGroup[],
): readonly ChartGroupInput[] {
  const axisElements = elementChildren(plotArea).filter((child) =>
    elementNamespaceUri(child) === CHART_NAMESPACE && AXIS_ELEMENTS.has(child.localName));
  const axes = axisElements.map((axis): ParsedAxis => ({
    id: readRequiredUnsignedChild(xml, axis, 'axId', `${axis.localName} axis id`),
    crossId: readRequiredUnsignedChild(xml, axis, 'crossAx', `${axis.localName} cross axis id`),
  }));
  requireUnique(axes.map(({ id }) => id), 'Chart axis ids');
  const axesById = new Map(axes.map((axis) => [axis.id, axis]));
  for (const axis of axes) {
    if (axis.id === axis.crossId || !axesById.has(axis.crossId)) {
      unsupported(`Chart axis ${axis.id} has a dangling cross-axis reference`);
    }
  }
  const usedAxes = new Set(groups.flatMap(({ axisIds }) => axisIds));
  if (axes.some(({ id }) => !usedAxes.has(id))) unsupported('Chart contains an unreferenced axis');
  for (const group of groups) {
    for (const id of group.axisIds) {
      const axis = axesById.get(id);
      if (!axis) unsupported(`${group.type} chart has a dangling axis reference`);
      if (!group.axisIds.includes(axis.crossId)) {
        unsupported(`${group.type} chart axes do not cross within the group`);
      }
    }
  }

  const axisSets: string[] = [];
  return Object.freeze(groups.map((group): ChartGroupInput => {
    if (group.axisIds.length === 0) return { type: group.type, series: group.series };
    const key = [...group.axisIds].sort((left, right) => left - right).join(',');
    let index = axisSets.indexOf(key);
    if (index < 0) {
      axisSets.push(key);
      index = axisSets.length - 1;
    }
    if (index > 1) unsupported('Chart uses more than primary and secondary axis sets');
    return {
      type: group.type,
      series: group.series,
      axis: index === 0 ? 'primary' : 'secondary',
    };
  }));
}

function readWorkbookPartUri(
  pkg: OpcPackage,
  chartPartUri: string,
  externalData: XmlElement,
): string {
  const references = attributesByExpandedName(externalData, RELATIONSHIP_NAMESPACE, 'id');
  if (references.length > 1) ambiguous('externalData has multiple relationship ids');
  if (references.length === 0) unsupported('externalData has no relationship id');
  const packageRelationships = pkg.relationships(chartPartUri).filter(
    ({ type }) => type === PACKAGE_RELATIONSHIP,
  );
  if (packageRelationships.length !== 1) {
    unsupported('Chart must have exactly one workbook package relationship');
  }
  const relationship = packageRelationships[0]!;
  if (
    relationship.id !== references[0]!.value
    || relationship.targetMode !== 'Internal'
    || !relationship.resolvedTarget
  ) {
    unsupported('Chart workbook relationship is invalid');
  }
  const workbook = pkg.getPart(relationship.resolvedTarget);
  if (!workbook || workbook.contentType !== XLSX_CONTENT_TYPE) {
    unsupported('Chart workbook relationship target is not an XLSX part');
  }
  return relationship.resolvedTarget;
}

function readRequiredIndex(
  xml: LosslessXmlDocument,
  parent: XmlElement,
  localName: string,
): number {
  const element = oneChartChild(
    parent,
    localName,
    `Chart series must contain exactly one ${localName}`,
  );
  return readUnsignedVal(xml, element, `Chart series ${localName}`, false);
}

function readRequiredUnsignedChild(
  xml: LosslessXmlDocument,
  parent: XmlElement,
  localName: string,
  context: string,
): number {
  return readUnsignedVal(
    xml,
    oneChartChild(parent, localName, `${context} must occur exactly once`),
    context,
    true,
  );
}

function readUnsignedVal(
  xml: LosslessXmlDocument,
  element: XmlElement,
  context: string,
  positive: boolean,
): number {
  return readUnsignedAttribute(xml, element, 'val', context, positive);
}

function readUnsignedAttribute(
  _xml: LosslessXmlDocument,
  element: XmlElement,
  name: string,
  context: string,
  positive: boolean,
): number {
  const attributes = element.attributes.filter((attribute) => attribute.name === name);
  if (attributes.length > 1) ambiguous(`${context} has multiple ${name} attributes`);
  if (attributes.length === 0 || !CANONICAL_INDEX_PATTERN.test(attributes[0]!.value)) {
    unsupported(`${context} must be a canonical unsigned integer`);
  }
  const value = Number(attributes[0]!.value);
  if (!Number.isSafeInteger(value) || (positive && value === 0)) {
    unsupported(`${context} is outside the supported range`);
  }
  return value;
}

function readDecimal(value: string, context: string): number {
  if (!DECIMAL_PATTERN.test(value)) unsupported(`${context} is not a finite decimal`);
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) unsupported(`${context} is not finite`);
  return Object.is(parsed, -0) ? 0 : parsed;
}

function requireUnique(values: readonly number[], context: string): void {
  if (new Set(values).size !== values.length) ambiguous(`${context} must be unique`);
}

function oneChartChild(parent: XmlElement, localName: string, message: string): XmlElement {
  const children = chartChildren(parent, localName);
  if (children.length > 1) ambiguous(message);
  if (children.length === 0) unsupported(message);
  return children[0]!;
}

function chartChildren(parent: XmlElement, localName: string): XmlElement[] {
  return elementChildren(parent).filter((child) =>
    child.localName === localName && elementNamespaceUri(child) === CHART_NAMESPACE);
}

function elementChildren(parent: XmlElement): XmlElement[] {
  return parent.children.filter((child): child is XmlElement => child.type === 'element');
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

function attributeNamespaceUri(
  element: XmlElement,
  attribute: XmlAttribute,
): string | undefined {
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

function stateError(
  status: 'unsupported' | 'ambiguous',
  reason: string,
): Readonly<ChartState> {
  return Object.freeze({ status, reason });
}

function unsupported(message: string): never {
  throw new ChartStateError('unsupported', message);
}

function ambiguous(message: string): never {
  throw new ChartStateError('ambiguous', message);
}
