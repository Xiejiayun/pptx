import { LosslessXmlDocument, type XmlAttribute, type XmlElement } from '@pptx/lossless-xml';
import type { OpcPackage } from '@pptx/opc';
import type {
  ChartAreaOptions,
  ChartAxisOptions,
  ChartCategories,
  ChartDataLabelOptions,
  ChartDataTableOptions,
  ChartDefinitionInput,
  ChartFontOptions,
  ChartGroupInput,
  ChartGroupOptions,
  ChartLegendOptions,
  ChartMarkerOptions,
  ChartOptions,
  ChartSeriesInput,
  ChartSeriesOptions,
  ChartState,
  ChartTitleOptions,
  ChartType,
} from './chart.js';
import { normalizeChartDefinition } from './chart-definition.internal.js';
import { readSimpleFillChoice, type SimpleFill } from './simple-fill.internal.js';
import { readSimpleLine, type NormalizedSimpleLine } from './simple-line.internal.js';

const CHART_NAMESPACE = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const MODERN_CHART_NAMESPACE = 'http://schemas.microsoft.com/office/drawing/2014/chartex';
const DRAWING_NAMESPACE = 'http://schemas.openxmlformats.org/drawingml/2006/main';
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
const LEGEND_POSITIONS = Object.freeze({
  b: 'bottom',
  l: 'left',
  r: 'right',
  t: 'top',
  tr: 'topRight',
} as const);
const AXIS_POSITIONS = Object.freeze({
  b: 'bottom',
  l: 'left',
  r: 'right',
  t: 'top',
} as const);
const LABEL_POSITIONS = Object.freeze({
  high: 'high',
  low: 'low',
  nextTo: 'nextTo',
  none: 'none',
} as const);

interface ParsedGroup {
  readonly type: ChartType;
  readonly series: readonly ChartSeriesInput[];
  readonly axisIds: readonly number[];
  readonly options?: ChartGroupOptions;
}

interface ParsedAxis {
  readonly id: number;
  readonly crossId: number;
  readonly element: XmlElement;
}

interface AxisAssignment {
  readonly groups: readonly ChartGroupInput[];
  readonly axisSets: readonly (readonly number[])[];
  readonly axesById: ReadonlyMap<number, ParsedAxis>;
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
  const assignment = assignAxes(xml, plotArea, parsedGroups);
  const options = readRootChartOptions(
    xml,
    root,
    chart,
    plotArea,
    groupElements,
    assignment,
    parsedGroups[0]!.type,
  );
  let definition;
  try {
    definition = normalizeChartDefinition({ groups: assignment.groups, options });
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
  const orderedSeries = indexed
    .sort((left, right) => left.order - right.order)
    .map(({ element }) => element);
  const series = orderedSeries.map((element) => readSeries(xml, element, type));
  const axisIds = chartChildren(group, 'axId').map((element) =>
    readUnsignedVal(xml, element, `${group.localName} axis reference`, true));
  requireUnique(axisIds, `${group.localName} axis references`);
  const expectedAxes = type === 'pie' || type === 'doughnut'
    ? [0]
    : type === 'bar3D'
      ? [2, 3]
      : type === 'area' || type === 'bar' || type === 'line' || type === 'radar'
        ? [2, 3]
        : [2];
  if (!expectedAxes.includes(axisIds.length)) {
    unsupported(`${group.localName} has an unsupported axis reference count`);
  }
  const options = readGroupOptions(xml, group, orderedSeries, type);
  return {
    type,
    series,
    axisIds: Object.freeze(axisIds),
    ...(options === undefined ? {} : { options }),
  };
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
): AxisAssignment {
  const axisElements = elementChildren(plotArea).filter((child) =>
    elementNamespaceUri(child) === CHART_NAMESPACE && AXIS_ELEMENTS.has(child.localName));
  const axes = axisElements.map((axis): ParsedAxis => ({
    id: readRequiredUnsignedChild(xml, axis, 'axId', `${axis.localName} axis id`),
    crossId: readRequiredUnsignedChild(xml, axis, 'crossAx', `${axis.localName} cross axis id`),
    element: axis,
  }));
  requireUnique(axes.map(({ id }) => id), 'Chart axis ids');
  const axesById = new Map(axes.map((axis) => [axis.id, axis]));
  const normalizedGroups = groups.map((group) =>
    normalizePptxGenJsTrailingAxisReference(group, axesById));
  for (const group of normalizedGroups) {
    const expected = group.type === 'pie' || group.type === 'doughnut'
      ? 0
      : group.type === 'bar3D'
        ? 3
        : 2;
    if (group.axisIds.length !== expected) {
      unsupported(`${group.type} chart has an unsupported axis reference count`);
    }
  }
  for (const axis of axes) {
    if (axis.id === axis.crossId || !axesById.has(axis.crossId)) {
      unsupported(`Chart axis ${axis.id} has a dangling cross-axis reference`);
    }
  }
  const usedAxes = new Set(normalizedGroups.flatMap(({ axisIds }) => axisIds));
  if (axes.some(({ id }) => !usedAxes.has(id))) unsupported('Chart contains an unreferenced axis');
  for (const group of normalizedGroups) {
    for (const id of group.axisIds) {
      const axis = axesById.get(id);
      if (!axis) unsupported(`${group.type} chart has a dangling axis reference`);
      if (!group.axisIds.includes(axis.crossId)) {
        unsupported(`${group.type} chart axes do not cross within the group`);
      }
    }
  }

  const axisSets: string[] = [];
  const assignedGroups = Object.freeze(normalizedGroups.map((group): ChartGroupInput => {
    if (group.axisIds.length === 0) {
      return {
        type: group.type,
        series: group.series,
        ...(group.options === undefined ? {} : { options: group.options }),
      } as ChartGroupInput;
    }
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
      ...(group.options === undefined ? {} : { options: group.options }),
    } as ChartGroupInput;
  }));
  return Object.freeze({
    groups: assignedGroups,
    axisSets: Object.freeze(axisSets.map((key) =>
      Object.freeze(key.split(',').map(Number)))),
    axesById,
  });
}

function normalizePptxGenJsTrailingAxisReference(
  group: ParsedGroup,
  axesById: ReadonlyMap<number, ParsedAxis>,
): ParsedGroup {
  if (
    !['area', 'bar', 'line', 'radar'].includes(group.type)
    || group.axisIds.length !== 3
    || !axesById.has(group.axisIds[0]!)
    || !axesById.has(group.axisIds[1]!)
    || axesById.has(group.axisIds[2]!)
  ) {
    return group;
  }
  return {
    ...group,
    axisIds: Object.freeze(group.axisIds.slice(0, 2)),
  };
}

function readRootChartOptions(
  xml: LosslessXmlDocument,
  root: XmlElement,
  chart: XmlElement,
  plotArea: XmlElement,
  groupElements: readonly XmlElement[],
  assignment: AxisAssignment,
  firstType: ChartType,
): ChartOptions {
  const result: Record<string, unknown> = {};
  const language = readChildStringValue(root, 'lang');
  if (language !== undefined) result.language = language;
  const style = readChildNumberValue(root, 'style');
  if (style !== undefined) result.style = style;
  if (readChildBooleanValue(root, 'roundedCorners') === true) result.roundedCorners = true;

  const title = optionalChartChild(chart, 'title');
  if (title) result.title = readTitleOptions(xml, title);
  const legend = optionalChartChild(chart, 'legend');
  if (legend) result.legend = readLegendOptions(legend);
  const blanks = readChildStringValue(chart, 'dispBlanksAs');
  if (blanks !== undefined && blanks !== 'gap') result.displayBlanksAs = blanks;

  const chartArea = optionalChartChild(root, 'spPr');
  const chartAreaOptions = chartArea ? readAreaOptions(chartArea) : undefined;
  if (
    chartAreaOptions
    && !(chartAreaOptions.fill?.kind === 'none' && chartAreaOptions.line?.kind === 'none')
  ) result.chartArea = chartAreaOptions;
  const plotAreaShape = optionalChartChild(plotArea, 'spPr');
  const plotAreaOptions = plotAreaShape ? readAreaOptions(plotAreaShape) : undefined;
  if (plotAreaOptions) result.plotArea = plotAreaOptions;

  const primary = assignment.axisSets[0];
  if (primary) {
    const { category, value } = resolveAxisRoles(assignment, primary, firstType, false);
    const categoryOptions = category
      ? readAxisOptions(xml, category, 'bottom', firstType === 'scatter' || firstType === 'bubble')
      : undefined;
    const valueOptions = value ? readAxisOptions(xml, value, 'left', false) : undefined;
    if (categoryOptions) result.categoryAxis = categoryOptions;
    if (valueOptions) result.valueAxis = valueOptions;
  }
  const secondary = assignment.axisSets[1];
  if (secondary) {
    const { category, value } = resolveAxisRoles(assignment, secondary, firstType, true);
    const categoryOptions = category
      ? readAxisOptions(xml, category, 'top', firstType === 'scatter' || firstType === 'bubble')
      : undefined;
    const valueOptions = value ? readAxisOptions(xml, value, 'right', false) : undefined;
    if (categoryOptions) result.secondaryCategoryAxis = categoryOptions;
    if (valueOptions) result.secondaryValueAxis = valueOptions;
  }

  const dataTable = optionalChartChild(plotArea, 'dTable');
  if (dataTable) result.dataTable = readDataTableOptions(xml, dataTable, groupElements);
  const view3D = optionalChartChild(chart, 'view3D');
  if (view3D) readView3DOptions(view3D, firstType, result);
  return result as ChartOptions;
}

function resolveAxisRoles(
  assignment: AxisAssignment,
  ids: readonly number[],
  type: ChartType,
  secondary: boolean,
): {
  readonly category: XmlElement | undefined;
  readonly value: XmlElement | undefined;
} {
  const elements = ids.map((id) => assignment.axesById.get(id)?.element).filter(
    (element): element is XmlElement => element !== undefined,
  );
  if (type !== 'scatter' && type !== 'bubble') {
    return {
      category: elements.find(({ localName }) => localName === 'catAx' || localName === 'dateAx'),
      value: elements.find(({ localName }) => localName === 'valAx'),
    };
  }
  const horizontalPosition = secondary ? 't' : 'b';
  const category = elements.find((element) =>
    readChildStringValue(element, 'axPos') === horizontalPosition) ?? elements[0];
  return { category, value: elements.find((element) => element !== category) };
}

function readGroupOptions(
  xml: LosslessXmlDocument,
  group: XmlElement,
  seriesElements: readonly XmlElement[],
  type: ChartType,
): ChartGroupOptions | undefined {
  const result: Record<string, unknown> = {};
  const varyColors = readChildBooleanValue(group, 'varyColors');
  const defaultVaryColors = type === 'pie' || type === 'doughnut';
  if (varyColors !== undefined && varyColors !== defaultVaryColors) result.varyColors = varyColors;

  const dataLabels = optionalChartChild(group, 'dLbls');
  if (dataLabels) result.dataLabels = readDataLabelOptions(dataLabels);
  const seriesOptions = seriesElements.map((series) => readSeriesOptions(series));
  if (seriesOptions.some((options) => Object.keys(options).length > 0)) {
    result.series = seriesOptions;
  }

  switch (type) {
    case 'area':
    case 'line': {
      const grouping = readChildStringValue(group, 'grouping');
      if (grouping !== undefined && grouping !== 'standard') result.grouping = grouping;
      if (type === 'line' && allSeriesBoolean(seriesElements, 'smooth', true)) result.smooth = true;
      break;
    }
    case 'bar':
    case 'bar3D': {
      const direction = readChildStringValue(group, 'barDir');
      if (direction === 'bar') result.direction = 'bar';
      const grouping = readChildStringValue(group, 'grouping');
      const defaultGrouping = type === 'bar3D' ? 'standard' : 'clustered';
      if (grouping !== undefined && grouping !== defaultGrouping) result.grouping = grouping;
      const gapWidth = readChildNumberValue(group, 'gapWidth');
      if (gapWidth !== undefined && gapWidth !== 150) result.gapWidth = gapWidth;
      if (type === 'bar') {
        const overlap = readChildNumberValue(group, 'overlap');
        if (overlap !== undefined && overlap !== 0) result.overlap = overlap;
      } else {
        const gapDepth = readChildNumberValue(group, 'gapDepth');
        if (gapDepth !== undefined && gapDepth !== 150) result.gapDepth = gapDepth;
      }
      break;
    }
    case 'bubble': {
      const scale = readChildNumberValue(group, 'bubbleScale');
      if (scale !== undefined && scale !== 100) result.scale = scale;
      if (readChildBooleanValue(group, 'showNegBubbles') === true) {
        result.showNegativeBubbles = true;
      }
      const represents = readChildStringValue(group, 'sizeRepresents');
      if (represents !== undefined && represents !== 'area') result.sizeRepresents = represents;
      break;
    }
    case 'doughnut': {
      const angle = readChildNumberValue(group, 'firstSliceAng');
      if (angle !== undefined && angle !== 0) result.firstSliceAngle = angle;
      const hole = readChildNumberValue(group, 'holeSize');
      if (hole !== undefined && hole !== 50) result.holeSize = hole;
      break;
    }
    case 'pie': {
      const angle = readChildNumberValue(group, 'firstSliceAng');
      if (angle !== undefined && angle !== 0) result.firstSliceAngle = angle;
      break;
    }
    case 'radar': {
      const style = readChildStringValue(group, 'radarStyle');
      if (style !== undefined && style !== 'standard') result.style = style;
      break;
    }
    case 'scatter': {
      const style = readChildStringValue(group, 'scatterStyle');
      if (style !== undefined && style !== 'lineMarker') result.style = style;
      if (allSeriesBoolean(seriesElements, 'smooth', true)) result.smooth = true;
      break;
    }
  }
  return Object.keys(result).length === 0 ? undefined : result as ChartGroupOptions;
}

function readSeriesOptions(series: XmlElement): ChartSeriesOptions {
  const result: Record<string, unknown> = {};
  const shape = optionalChartChild(series, 'spPr');
  if (shape) Object.assign(result, readAreaOptions(shape));
  const marker = optionalChartChild(series, 'marker');
  if (marker) result.marker = readMarkerOptions(marker);
  return result as ChartSeriesOptions;
}

function readMarkerOptions(marker: XmlElement): ChartMarkerOptions {
  const result: Record<string, unknown> = {};
  const shape = readChildStringValue(marker, 'symbol');
  if (shape !== undefined && shape !== 'circle') result.shape = shape;
  const size = readChildNumberValue(marker, 'size');
  if (size !== undefined && size !== 5) result.size = size;
  const properties = optionalChartChild(marker, 'spPr');
  if (properties) Object.assign(result, readAreaOptions(properties));
  return result as ChartMarkerOptions;
}

function readTitleOptions(
  xml: LosslessXmlDocument,
  title: XmlElement,
): ChartTitleOptions {
  const result: Record<string, unknown> = { ...readFontOptions(title, true) };
  const text = drawingDescendants(title, 't').map((element) => xml.text(element)).join('');
  if (text.length > 0) result.text = text;
  if (readChildBooleanValue(title, 'overlay') === true) result.overlay = true;
  const body = drawingDescendants(title, 'bodyPr')[0];
  const rotation = body ? readNumberAttribute(body, 'rot') : undefined;
  if (rotation !== undefined && rotation !== 0) result.rotation = rotation / 60_000;
  const layout = optionalChartChild(title, 'layout');
  const manual = layout ? optionalChartChild(layout, 'manualLayout') : undefined;
  const x = manual ? readChildNumberValue(manual, 'x') : undefined;
  const y = manual ? readChildNumberValue(manual, 'y') : undefined;
  if (x !== undefined && y !== undefined) result.position = { x, y };
  return result as ChartTitleOptions;
}

function readLegendOptions(legend: XmlElement): ChartLegendOptions {
  const result: Record<string, unknown> = { ...readFontOptions(legend, false) };
  const position = readChildStringValue(legend, 'legendPos');
  const mappedPosition = position === undefined
    ? undefined
    : (LEGEND_POSITIONS as Readonly<Record<string, ChartLegendOptions['position']>>)[position];
  if (mappedPosition !== undefined && mappedPosition !== 'right') result.position = mappedPosition;
  if (readChildBooleanValue(legend, 'overlay') === true) result.overlay = true;
  return result as ChartLegendOptions;
}

function readAxisOptions(
  xml: LosslessXmlDocument,
  axis: XmlElement,
  defaultPosition: 'bottom' | 'left' | 'right' | 'top',
  horizontalValueAxis: boolean,
): ChartAxisOptions | undefined {
  const result: Record<string, unknown> = { ...readFontOptions(axis, false) };
  if (readChildBooleanValue(axis, 'delete') === true) result.visible = false;
  const rawPosition = readChildStringValue(axis, 'axPos');
  const position = rawPosition === undefined
    ? undefined
    : (AXIS_POSITIONS as Readonly<Record<string, ChartAxisOptions['position']>>)[rawPosition];
  if (position !== undefined && position !== defaultPosition) result.position = position;
  const title = optionalChartChild(axis, 'title');
  if (title) result.title = readTitleOptions(xml, title);

  const scaling = optionalChartChild(axis, 'scaling');
  if (scaling) {
    const minimum = readChildNumberValue(scaling, 'min');
    const maximum = readChildNumberValue(scaling, 'max');
    const logBase = readChildNumberValue(scaling, 'logBase');
    const orientation = readChildStringValue(scaling, 'orientation');
    if (minimum !== undefined) result.minimum = minimum;
    if (maximum !== undefined) result.maximum = maximum;
    if (logBase !== undefined) result.logarithmicBase = logBase;
    if (orientation !== undefined && orientation !== 'minMax') result.orientation = orientation;
  }
  const majorUnit = readChildNumberValue(axis, 'majorUnit');
  const minorUnit = readChildNumberValue(axis, 'minorUnit');
  if (majorUnit !== undefined) result.majorUnit = majorUnit;
  if (minorUnit !== undefined) result.minorUnit = minorUnit;
  const format = optionalChartChild(axis, 'numFmt');
  const formatCode = format ? readStringAttribute(format, 'formatCode') : undefined;
  if (formatCode !== undefined && formatCode !== 'General') result.numberFormat = formatCode;
  const labelPosition = readChildStringValue(axis, 'tickLblPos');
  const mappedLabel = labelPosition === undefined
    ? undefined
    : (LABEL_POSITIONS as Readonly<Record<string, ChartAxisOptions['labelPosition']>>)[labelPosition];
  if (mappedLabel !== undefined && mappedLabel !== 'nextTo') result.labelPosition = mappedLabel;
  const textProperties = optionalChartChild(axis, 'txPr');
  const body = textProperties ? drawingDescendants(textProperties, 'bodyPr')[0] : undefined;
  const rotation = body ? readNumberAttribute(body, 'rot') : undefined;
  if (rotation !== undefined && rotation !== 0) result.labelRotation = rotation / 60_000;
  const shape = optionalChartChild(axis, 'spPr');
  if (shape) {
    const line = readLineFromProperties(shape);
    if (line) result.line = line;
  }
  const majorGridLine = optionalChartChild(axis, 'majorGridlines');
  const majorGridProperties = majorGridLine ? optionalChartChild(majorGridLine, 'spPr') : undefined;
  const majorGrid = majorGridProperties ? readLineFromProperties(majorGridProperties) : undefined;
  if (majorGrid) result.majorGridLine = majorGrid;
  const minorGridLine = optionalChartChild(axis, 'minorGridlines');
  const minorGridProperties = minorGridLine ? optionalChartChild(minorGridLine, 'spPr') : undefined;
  const minorGrid = minorGridProperties ? readLineFromProperties(minorGridProperties) : undefined;
  if (minorGrid) result.minorGridLine = minorGrid;
  const defaultMajorTick = horizontalValueAxis ? 'none' : 'out';
  const majorTick = readChildStringValue(axis, 'majorTickMark');
  const minorTick = readChildStringValue(axis, 'minorTickMark');
  const mappedMajor = mapTickMark(majorTick);
  const mappedMinor = mapTickMark(minorTick);
  if (mappedMajor !== undefined && majorTick !== defaultMajorTick) result.majorTickMark = mappedMajor;
  if (mappedMinor !== undefined && minorTick !== 'none') result.minorTickMark = mappedMinor;
  return Object.keys(result).length === 0 ? undefined : result as ChartAxisOptions;
}

function readDataLabelOptions(labels: XmlElement): ChartDataLabelOptions {
  const result: Record<string, unknown> = { ...readFontOptions(labels, false) };
  const flags: readonly [string, keyof ChartDataLabelOptions][] = [
    ['showVal', 'showValue'],
    ['showCatName', 'showCategoryName'],
    ['showSerName', 'showSeriesName'],
    ['showPercent', 'showPercent'],
    ['showBubbleSize', 'showBubbleSize'],
    ['showLeaderLines', 'showLeaderLines'],
  ];
  for (const [elementName, property] of flags) {
    if (readChildBooleanValue(labels, elementName) === true) result[property] = true;
  }
  const position = readChildStringValue(labels, 'dLblPos');
  const mappedPosition = position === undefined
    ? undefined
    : ({
        bestFit: 'bestFit',
        b: 'bottom',
        ctr: 'center',
        inBase: 'insideBase',
        inEnd: 'insideEnd',
        l: 'left',
        outEnd: 'outsideEnd',
        r: 'right',
        t: 'top',
      } as Readonly<Record<string, ChartDataLabelOptions['position']>>)[position];
  if (mappedPosition !== undefined) result.position = mappedPosition;
  const format = optionalChartChild(labels, 'numFmt');
  const formatCode = format ? readStringAttribute(format, 'formatCode') : undefined;
  if (formatCode !== undefined) result.numberFormat = formatCode;
  return result as ChartDataLabelOptions;
}

function readDataTableOptions(
  xml: LosslessXmlDocument,
  table: XmlElement,
  groupElements: readonly XmlElement[],
): ChartDataTableOptions {
  const result: Record<string, unknown> = { ...readFontOptions(table, false) };
  if (readChildBooleanValue(table, 'showHorzBorder') === false) result.showHorizontalBorder = false;
  if (readChildBooleanValue(table, 'showVertBorder') === false) result.showVerticalBorder = false;
  if (readChildBooleanValue(table, 'showOutline') === false) result.showOutline = false;
  if (readChildBooleanValue(table, 'showKeys') === false) result.showLegendKeys = false;
  const formats = groupElements.flatMap((group) =>
    chartChildren(group, 'ser').map((series) => readSeriesValueFormat(xml, series)).filter(
      (value): value is string => value !== undefined,
    ));
  if (formats.length > 0 && formats.every((value) => value === formats[0]) && formats[0] !== 'General') {
    result.numberFormat = formats[0];
  }
  return result as ChartDataTableOptions;
}

function readView3DOptions(
  view: XmlElement,
  type: ChartType,
  result: Record<string, unknown>,
): void {
  const rotationX = readChildNumberValue(view, 'rotX');
  const rotationY = readChildNumberValue(view, 'rotY');
  const rightAngleAxes = readChildBooleanValue(view, 'rAngAx');
  const perspective = readChildNumberValue(view, 'perspective');
  const defaults = type === 'bar3D'
    ? { rotationX: 15, rotationY: 20, rightAngleAxes: true, perspective: 30 }
    : {};
  if (rotationX !== undefined && rotationX !== defaults.rotationX) result.rotationX = rotationX;
  if (rotationY !== undefined && rotationY !== defaults.rotationY) result.rotationY = rotationY;
  if (rightAngleAxes !== undefined && rightAngleAxes !== defaults.rightAngleAxes) {
    result.rightAngleAxes = rightAngleAxes;
  }
  if (perspective !== undefined && perspective !== defaults.perspective) result.perspective = perspective;
}

function readAreaOptions(properties: XmlElement): ChartAreaOptions | undefined {
  const result: Record<string, unknown> = {};
  const fill = readFillFromProperties(properties);
  const line = readLineFromProperties(properties);
  if (fill) result.fill = fill;
  if (line) result.line = line;
  return Object.keys(result).length === 0 ? undefined : result as ChartAreaOptions;
}

function readFillFromProperties(properties: XmlElement): SimpleFill | undefined {
  const fills = elementChildren(properties).filter((child) =>
    elementNamespaceUri(child) === DRAWING_NAMESPACE
      && (child.localName === 'noFill' || child.localName === 'solidFill'));
  return fills.length === 1 ? readSimpleFillChoice(fills[0]!, 'a:') : undefined;
}

function readLineFromProperties(properties: XmlElement): NormalizedSimpleLine | undefined {
  const lines = elementChildren(properties).filter((child) =>
    elementNamespaceUri(child) === DRAWING_NAMESPACE && child.localName === 'ln');
  return lines.length === 1 ? readSimpleLine(lines[0]!, 'a:') : undefined;
}

function readFontOptions(parent: XmlElement, preferRun: boolean): ChartFontOptions {
  const runs = drawingDescendants(parent, preferRun ? 'rPr' : 'defRPr');
  const fallback = drawingDescendants(parent, preferRun ? 'defRPr' : 'rPr');
  const properties = runs[0] ?? fallback[0];
  if (!properties) return {};
  const result: Record<string, unknown> = {};
  const size = readNumberAttribute(properties, 'sz');
  if (size !== undefined && size > 0) result.size = size / 100;
  if (readBooleanAttribute(properties, 'b') === true) result.bold = true;
  if (readBooleanAttribute(properties, 'i') === true) result.italic = true;
  const fill = elementChildren(properties).find((child) =>
    elementNamespaceUri(child) === DRAWING_NAMESPACE && child.localName === 'solidFill');
  const parsedFill = fill ? readSimpleFillChoice(fill, 'a:') : undefined;
  if (parsedFill?.kind === 'solid') result.color = parsedFill.color;
  const latin = elementChildren(properties).find((child) =>
    elementNamespaceUri(child) === DRAWING_NAMESPACE && child.localName === 'latin');
  const face = latin ? readStringAttribute(latin, 'typeface') : undefined;
  if (face !== undefined && face.length > 0) result.face = face;
  return result as ChartFontOptions;
}

function readSeriesValueFormat(xml: LosslessXmlDocument, series: XmlElement): string | undefined {
  const container = optionalChartChild(series, 'val') ?? optionalChartChild(series, 'yVal');
  const reference = container ? optionalChartChild(container, 'numRef') : undefined;
  const cache = reference ? optionalChartChild(reference, 'numCache') : undefined;
  const format = cache ? optionalChartChild(cache, 'formatCode') : undefined;
  return format ? xml.text(format) : undefined;
}

function allSeriesBoolean(
  series: readonly XmlElement[],
  localName: string,
  value: boolean,
): boolean {
  return series.length > 0 && series.every((entry) => readChildBooleanValue(entry, localName) === value);
}

function mapTickMark(value: string | undefined): ChartAxisOptions['majorTickMark'] | undefined {
  if (value === undefined) return undefined;
  return ({ cross: 'cross', in: 'inside', none: 'none', out: 'outside' } as const)[value as 'cross'];
}

function optionalChartChild(parent: XmlElement, localName: string): XmlElement | undefined {
  const children = chartChildren(parent, localName);
  return children.length === 1 ? children[0] : undefined;
}

function readChildStringValue(parent: XmlElement, localName: string): string | undefined {
  const child = optionalChartChild(parent, localName);
  return child ? readStringAttribute(child, 'val') : undefined;
}

function readChildNumberValue(parent: XmlElement, localName: string): number | undefined {
  const value = readChildStringValue(parent, localName);
  if (value === undefined || !DECIMAL_PATTERN.test(value)) return undefined;
  const number = Number(value);
  return Number.isFinite(number) ? number : undefined;
}

function readChildBooleanValue(parent: XmlElement, localName: string): boolean | undefined {
  const value = readChildStringValue(parent, localName);
  return value === '1' || value === 'true'
    ? true
    : value === '0' || value === 'false'
      ? false
      : undefined;
}

function readStringAttribute(element: XmlElement, name: string): string | undefined {
  const attributes = element.attributes.filter((attribute) => attribute.name === name);
  return attributes.length === 1 ? attributes[0]!.value : undefined;
}

function readNumberAttribute(element: XmlElement, name: string): number | undefined {
  const value = readStringAttribute(element, name);
  if (value === undefined || !/^[+-]?\d+$/.test(value)) return undefined;
  const number = Number(value);
  return Number.isSafeInteger(number) ? number : undefined;
}

function readBooleanAttribute(element: XmlElement, name: string): boolean | undefined {
  const value = readStringAttribute(element, name);
  return value === '1' || value === 'true'
    ? true
    : value === '0' || value === 'false'
      ? false
      : undefined;
}

function drawingDescendants(parent: XmlElement, localName: string): XmlElement[] {
  const result: XmlElement[] = [];
  const visit = (element: XmlElement): void => {
    for (const child of elementChildren(element)) {
      if (elementNamespaceUri(child) === DRAWING_NAMESPACE && child.localName === localName) {
        result.push(child);
      }
      visit(child);
    }
  };
  visit(parent);
  return result;
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
