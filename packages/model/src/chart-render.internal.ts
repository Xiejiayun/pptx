import { escapeXmlAttribute, escapeXmlText } from '@pptx/lossless-xml';
import type {
  AddChartOptions,
  ChartAreaOptions,
  ChartAxisOptions,
  ChartCategories,
  ChartCategoryAxisOptions,
  ChartDataLabelOptions,
  ChartDataTableOptions,
  ChartDefinition,
  ChartFontOptions,
  ChartGroupOptions,
  ChartLegendOptions,
  ChartMarkerOptions,
  ChartSeries,
  ChartSeriesAxisOptions,
  ChartSeriesOptions,
  ChartTitleOptions,
  ChartType,
  ChartValueAxisOptions,
} from './chart.js';
import { normalizePlaceholderSelector } from './placeholder.internal.js';
import type { PlaceholderIdentity, PlaceholderSelector } from './placeholder.js';
import type {
  ChartWorkbookFormula,
  ChartWorkbookPlan,
} from './chart-workbook.internal.js';
import { renderSimpleFill, type SimpleFill } from './simple-fill.internal.js';
import {
  renderSimpleLine,
  type NormalizedSimpleLine,
} from './simple-line.internal.js';
import type { RichTextColor } from './text.js';
import { inches, type Emu, type OoxmlAngle } from './units.js';

const CHART_NAMESPACE = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const DRAWING_NAMESPACE = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CHART_GRAPHIC_DATA_URI = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const MAX_ROTATION = 21_600_000;
const EMU_PER_POINT = 12_700;
const OPTION_KEYS = new Set([
  'name',
  'altText',
  'placeholder',
  'x',
  'y',
  'width',
  'height',
  'rotation',
  'flipHorizontal',
  'flipVertical',
]);

export interface NormalizedAddChartOptions {
  readonly name: string;
  readonly altText?: string;
  readonly placeholder?: PlaceholderSelector;
  readonly x: Emu;
  readonly y: Emu;
  readonly width: Emu;
  readonly height: Emu;
  readonly rotation: OoxmlAngle;
  readonly flipHorizontal: boolean;
  readonly flipVertical: boolean;
}

export function normalizeAddChartOptions(
  value: AddChartOptions | undefined = undefined,
): Readonly<NormalizedAddChartOptions> {
  const options = readOptions(value);
  const name = readXmlString(options.name, 'Chart name', 'Chart');
  const altText = options.altText === undefined
    ? undefined
    : readXmlString(options.altText, 'Chart altText');
  const placeholder = options.placeholder === undefined
    ? undefined
    : normalizePlaceholderSelector(options.placeholder);
  const x = readCoordinate(options.x, inches(1), 'Chart x') as Emu;
  const y = readCoordinate(options.y, inches(1), 'Chart y') as Emu;
  const width = readCoordinate(options.width, inches(6), 'Chart width') as Emu;
  const height = readCoordinate(options.height, inches(4), 'Chart height') as Emu;
  if (width <= 0) throw new RangeError('Chart width must be greater than zero');
  if (height <= 0) throw new RangeError('Chart height must be greater than zero');
  const rotation = readCoordinate(options.rotation, 0, 'Chart rotation');
  if (rotation < -MAX_ROTATION || rotation > MAX_ROTATION) {
    throw new RangeError('Chart rotation must be between -21600000 and 21600000');
  }
  return Object.freeze({
    name,
    ...(altText === undefined ? {} : { altText }),
    ...(placeholder === undefined ? {} : { placeholder }),
    x,
    y,
    width,
    height,
    rotation: rotation as OoxmlAngle,
    flipHorizontal: readBoolean(options.flipHorizontal, false, 'Chart flipHorizontal'),
    flipVertical: readBoolean(options.flipVertical, false, 'Chart flipVertical'),
  });
}

export function renderChartPart(
  definition: Readonly<ChartDefinition>,
  formulas: Readonly<ChartWorkbookPlan['formulas']>,
  workbookRelationshipId: string,
): string {
  const allocation = allocateAxes(definition);
  const chartOptions = definition.options;
  let globalSeriesIndex = 0;
  const groupXml = definition.groups.map((group, groupIndex) => {
    const axisIds = (group.axis ?? 'primary') === 'secondary'
      ? allocation.secondary
      : allocation.primary;
    if (!axisIds) throw new Error(`Chart group ${groupIndex} has no allocated axis set`);
    const renderedSeries = group.series.map((series, seriesIndex) => {
      const formula = formulas.find((candidate) =>
        candidate.groupIndex === groupIndex && candidate.seriesIndex === seriesIndex);
      if (!formula) {
        throw new Error(`Chart group ${groupIndex} series ${seriesIndex} has no workbook formula plan`);
      }
      const index = globalSeriesIndex;
      globalSeriesIndex += 1;
      const seriesOptions = group.options?.series?.[seriesIndex];
      return group.type === 'scatter' || group.type === 'bubble'
        ? renderXySeries(
            series,
            formula,
            index,
            group.type,
            seriesOptions,
            group.options,
            chartOptions.colors,
            chartOptions.dataTable?.numberFormat,
          )
        : renderCategoricalSeries(
            series,
            formula,
            index,
            group.type,
            seriesOptions,
            group.options,
            chartOptions.colors,
            chartOptions.dataTable?.numberFormat,
          );
    }).join('');
    return group.type === 'scatter' || group.type === 'bubble'
      ? renderXyGroup(group.type, renderedSeries, axisIds, group.options)
      : renderCategoricalGroup(group.type, renderedSeries, axisIds, group.options);
  }).join('');
  const axes = renderAllocatedAxes(definition, allocation);
  const title = renderChartTitle(chartOptions.title, chartOptions.language);
  const legend = renderLegend(chartOptions.legend, chartOptions.language);
  const view3D = renderView3D(definition);
  const plotAreaShape = renderAreaShapeProperties(chartOptions.plotArea, false);
  const dataTable = renderDataTable(chartOptions.dataTable, chartOptions.language);
  const style = chartOptions.style === undefined ? '' : `<c:style val="${chartOptions.style}"/>`;
  const language = chartOptions.language === undefined
    ? ''
    : `<c:lang val="${escapeXmlAttribute(chartOptions.language)}"/>`;
  const roundedCorners = chartOptions.roundedCorners ? 1 : 0;
  const autoTitleDeleted = chartOptions.title?.visible === false || !chartOptions.title
    ? 1
    : 0;
  const blanks = chartOptions.displayBlanksAs ?? 'gap';
  const chartAreaShape = renderAreaShapeProperties(chartOptions.chartArea, true);
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + `<c:chartSpace xmlns:c="${CHART_NAMESPACE}" xmlns:a="${DRAWING_NAMESPACE}" `
    + `xmlns:r="${RELATIONSHIP_NAMESPACE}">`
    + `<c:date1904 val="0"/>${language}<c:roundedCorners val="${roundedCorners}"/>${style}`
    + `<c:chart>${title}<c:autoTitleDeleted val="${autoTitleDeleted}"/>${view3D}`
    + `<c:plotArea><c:layout/>${groupXml}${axes}${dataTable}${plotAreaShape}</c:plotArea>`
    + `${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="${blanks}"/></c:chart>`
    + chartAreaShape
    + `<c:externalData r:id="${escapeXmlAttribute(workbookRelationshipId)}">`
    + '<c:autoUpdate val="0"/></c:externalData></c:chartSpace>';
}

interface AxisAllocation {
  readonly primary: readonly number[];
  readonly secondary?: readonly number[];
}

function allocateAxes(definition: Readonly<ChartDefinition>): AxisAllocation {
  const firstType = definition.groups[0]!.type;
  if (firstType === 'pie' || firstType === 'doughnut') return { primary: [] };
  if (firstType === 'bar3D') return { primary: [10_000_001, 10_000_002, 10_000_003] };
  const secondary = definition.groups.some(({ axis }) => axis === 'secondary')
    ? [10_000_003, 10_000_004]
    : undefined;
  return {
    primary: [10_000_001, 10_000_002],
    ...(secondary === undefined ? {} : { secondary }),
  };
}

function renderAllocatedAxes(
  definition: Readonly<ChartDefinition>,
  allocation: AxisAllocation,
): string {
  const type = definition.groups[0]!.type;
  return renderAxes(
    type,
    allocation.primary,
    false,
    definition.options.categoryAxis,
    definition.options.valueAxis,
    definition.options.seriesAxis,
  ) + (allocation.secondary
    ? renderAxes(
        type,
        allocation.secondary,
        true,
        definition.options.secondaryCategoryAxis,
        definition.options.secondaryValueAxis,
        undefined,
      )
    : '');
}

export function renderChartGraphicFrame(
  shapeId: number,
  relationshipId: string,
  options: Readonly<NormalizedAddChartOptions>,
  placeholder?: Readonly<PlaceholderIdentity>,
): string {
  const transformAttributes = [
    options.rotation === 0 ? '' : ` rot="${options.rotation}"`,
    options.flipHorizontal ? ' flipH="1"' : '',
    options.flipVertical ? ' flipV="1"' : '',
  ].join('');
  const description = options.altText === undefined
    ? ''
    : ` descr="${escapeXmlAttribute(options.altText)}"`;
  const applicationProperties = placeholder === undefined
    ? '<p:nvPr/>'
    : `<p:nvPr><p:ph type="${placeholder.type}" idx="${placeholder.index}"/></p:nvPr>`;
  return `<p:graphicFrame xmlns:p="${PRESENTATION_NAMESPACE}" xmlns:a="${DRAWING_NAMESPACE}" `
    + `xmlns:c="${CHART_NAMESPACE}" xmlns:r="${RELATIONSHIP_NAMESPACE}">`
    + '<p:nvGraphicFramePr>'
    + `<p:cNvPr id="${shapeId}" name="${escapeXmlAttribute(options.name)}"${description}/>`
    + '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>'
    + applicationProperties
    + '</p:nvGraphicFramePr>'
    + `<p:xfrm${transformAttributes}><a:off x="${options.x}" y="${options.y}"/>`
    + `<a:ext cx="${options.width}" cy="${options.height}"/></p:xfrm>`
    + `<a:graphic><a:graphicData uri="${CHART_GRAPHIC_DATA_URI}">`
    + `<c:chart r:id="${escapeXmlAttribute(relationshipId)}"/>`
    + '</a:graphicData></a:graphic></p:graphicFrame>';
}

function renderCategoricalGroup(
  type: Exclude<ChartType, 'scatter' | 'bubble'>,
  series: string,
  axisIds: readonly number[],
  options: Readonly<ChartGroupOptions> | undefined,
): string {
  const axes = axisIds.map((id) => `<c:axId val="${id}"/>`).join('');
  const dataLabels = renderDataLabels(options?.dataLabels);
  const varyColors = groupOption<boolean>(options, 'varyColors')
    ?? (type === 'pie' || type === 'doughnut');
  switch (type) {
    case 'area':
      return `<c:areaChart><c:grouping val="${groupOption<string>(options, 'grouping') ?? 'standard'}"/>`
        + `<c:varyColors val="${varyColors ? 1 : 0}"/>${series}${dataLabels}${axes}</c:areaChart>`;
    case 'bar':
      return `<c:barChart><c:barDir val="${groupOption<string>(options, 'direction') === 'bar' ? 'bar' : 'col'}"/>`
        + `<c:grouping val="${groupOption<string>(options, 'grouping') ?? 'clustered'}"/>`
        + `<c:varyColors val="${varyColors ? 1 : 0}"/>${series}${dataLabels}`
        + `<c:gapWidth val="${groupOption<number>(options, 'gapWidth') ?? 150}"/>`
        + `<c:overlap val="${groupOption<number>(options, 'overlap') ?? 0}"/>${axes}</c:barChart>`;
    case 'bar3D':
      return `<c:bar3DChart><c:barDir val="${groupOption<string>(options, 'direction') === 'bar' ? 'bar' : 'col'}"/>`
        + `<c:grouping val="${groupOption<string>(options, 'grouping') ?? 'standard'}"/>`
        + `<c:varyColors val="${varyColors ? 1 : 0}"/>${series}${dataLabels}`
        + `<c:gapWidth val="${groupOption<number>(options, 'gapWidth') ?? 150}"/>`
        + `<c:gapDepth val="${groupOption<number>(options, 'gapDepth') ?? 150}"/>${axes}</c:bar3DChart>`;
    case 'doughnut':
      return `<c:doughnutChart><c:varyColors val="${varyColors ? 1 : 0}"/>${series}${dataLabels}`
        + `<c:firstSliceAng val="${groupOption<number>(options, 'firstSliceAngle') ?? 0}"/>`
        + `<c:holeSize val="${groupOption<number>(options, 'holeSize') ?? 50}"/></c:doughnutChart>`;
    case 'line':
      return `<c:lineChart><c:grouping val="${groupOption<string>(options, 'grouping') ?? 'standard'}"/>`
        + `<c:varyColors val="${varyColors ? 1 : 0}"/>${series}${dataLabels}${axes}</c:lineChart>`;
    case 'pie':
      return `<c:pieChart><c:varyColors val="${varyColors ? 1 : 0}"/>${series}${dataLabels}`
        + `<c:firstSliceAng val="${groupOption<number>(options, 'firstSliceAngle') ?? 0}"/></c:pieChart>`;
    case 'radar':
      return `<c:radarChart><c:radarStyle val="${groupOption<string>(options, 'style') ?? 'standard'}"/>`
        + `<c:varyColors val="${varyColors ? 1 : 0}"/>${series}${dataLabels}${axes}</c:radarChart>`;
  }
}

function renderCategoricalSeries(
  series: Readonly<ChartSeries>,
  formula: Readonly<ChartWorkbookFormula>,
  index: number,
  type: Exclude<ChartType, 'scatter' | 'bubble'>,
  options: Readonly<ChartSeriesOptions> | undefined,
  groupOptions: Readonly<ChartGroupOptions> | undefined,
  colors: readonly RichTextColor[] | undefined,
  valueNumberFormat: string | undefined,
): string {
  const style = renderSeriesShapeProperties(type, index, options, colors);
  const marker = type === 'line' || type === 'radar'
    ? renderMarker(options?.marker ?? groupOption<ChartMarkerOptions>(groupOptions, 'marker'))
    : '';
  const smooth = type === 'line'
    ? `<c:smooth val="${groupOption<boolean>(groupOptions, 'smooth') ? 1 : 0}"/>`
    : '';
  return '<c:ser>'
    + `<c:idx val="${index}"/><c:order val="${index}"/>`
    + renderSeriesName(series.name, formula.name)
    + style
    + marker
    + renderCategories(series.categories!, formula.categories!)
    + '<c:val><c:numRef>'
    + `<c:f>${escapeXmlText(formula.values)}</c:f>${renderNumericCache(series.values, valueNumberFormat)}`
    + `</c:numRef></c:val>${smooth}</c:ser>`;
}

function renderXyGroup(
  type: 'scatter' | 'bubble',
  series: string,
  axisIds: readonly number[],
  options: Readonly<ChartGroupOptions> | undefined,
): string {
  const axes = axisIds.map((id) => `<c:axId val="${id}"/>`).join('');
  const dataLabels = renderDataLabels(options?.dataLabels);
  const varyColors = groupOption<boolean>(options, 'varyColors') ?? false;
  if (type === 'scatter') {
    return `<c:scatterChart><c:scatterStyle val="${groupOption<string>(options, 'style') ?? 'lineMarker'}"/>`
      + `<c:varyColors val="${varyColors ? 1 : 0}"/>${series}${dataLabels}${axes}</c:scatterChart>`;
  }
  return `<c:bubbleChart><c:varyColors val="${varyColors ? 1 : 0}"/>${series}${dataLabels}`
    + `<c:bubbleScale val="${groupOption<number>(options, 'scale') ?? 100}"/>`
    + `<c:showNegBubbles val="${groupOption<boolean>(options, 'showNegativeBubbles') ? 1 : 0}"/>`
    + `<c:sizeRepresents val="${groupOption<string>(options, 'sizeRepresents') ?? 'area'}"/>`
    + `${axes}</c:bubbleChart>`;
}

function renderXySeries(
  series: Readonly<ChartSeries>,
  formula: Readonly<ChartWorkbookFormula>,
  index: number,
  type: 'scatter' | 'bubble',
  options: Readonly<ChartSeriesOptions> | undefined,
  groupOptions: Readonly<ChartGroupOptions> | undefined,
  colors: readonly RichTextColor[] | undefined,
  valueNumberFormat: string | undefined,
): string {
  if (!formula.xValues || (type === 'bubble' && !formula.sizes)) {
    throw new Error(`Chart ${type} series ${index} has an incomplete workbook formula plan`);
  }
  const style = renderSeriesShapeProperties(type, index, options, colors);
  const marker = type === 'scatter'
    ? renderMarker(options?.marker ?? groupOption<ChartMarkerOptions>(groupOptions, 'marker'))
    : '';
  const smooth = groupOption<boolean>(groupOptions, 'smooth') ?? false;
  return '<c:ser>'
    + `<c:idx val="${index}"/><c:order val="${index}"/>`
    + renderSeriesName(series.name, formula.name)
    + style
    + marker
    + '<c:xVal><c:numRef>'
    + `<c:f>${escapeXmlText(formula.xValues)}</c:f>${renderNumericCache(series.xValues!)}`
    + '</c:numRef></c:xVal>'
    + '<c:yVal><c:numRef>'
    + `<c:f>${escapeXmlText(formula.values)}</c:f>${renderNumericCache(series.values, valueNumberFormat)}`
    + '</c:numRef></c:yVal>'
    + (type === 'bubble'
      ? '<c:bubbleSize><c:numRef>'
        + `<c:f>${escapeXmlText(formula.sizes!)}</c:f>${renderNumericCache(series.sizes!)}`
        + '</c:numRef></c:bubbleSize><c:bubble3D val="0"/>'
      : `<c:smooth val="${smooth ? 1 : 0}"/>`)
    + '</c:ser>';
}

function renderSeriesName(name: string, formula: string): string {
  return '<c:tx><c:strRef>'
    + `<c:f>${escapeXmlText(formula)}</c:f><c:strCache><c:ptCount val="1"/>`
    + `<c:pt idx="0"><c:v>${escapeXmlText(name)}</c:v></c:pt>`
    + '</c:strCache></c:strRef></c:tx>';
}

function renderCategories(categories: ChartCategories, formulas: readonly string[]): string {
  const first = categories[0];
  if (Array.isArray(first)) {
    const levels = categories as readonly (readonly string[])[];
    const formula = rectangularCategoryFormula(formulas);
    return '<c:cat><c:multiLvlStrRef>'
      + `<c:f>${escapeXmlText(formula)}</c:f><c:multiLvlStrCache>`
      + `<c:ptCount val="${first.length}"/>`
      + levels.map((level) => `<c:lvl>${renderStringPoints(level)}</c:lvl>`).join('')
      + '</c:multiLvlStrCache></c:multiLvlStrRef></c:cat>';
  }
  const values = categories as readonly (string | number)[];
  if (values.every((value) => typeof value === 'number')) {
    return '<c:cat><c:numRef>'
      + `<c:f>${escapeXmlText(formulas[0]!)}</c:f>${renderNumericCache(values as readonly number[])}`
      + '</c:numRef></c:cat>';
  }
  return '<c:cat><c:strRef>'
    + `<c:f>${escapeXmlText(formulas[0]!)}</c:f><c:strCache>`
    + `<c:ptCount val="${values.length}"/>${renderStringPoints(values.map(String))}`
    + '</c:strCache></c:strRef></c:cat>';
}

function renderStringPoints(values: readonly string[]): string {
  return values.map((value, index) =>
    `<c:pt idx="${index}"><c:v>${escapeXmlText(value)}</c:v></c:pt>`).join('');
}

function renderNumericCache(values: readonly number[], formatCode = 'General'): string {
  return `<c:numCache><c:formatCode>${escapeXmlText(formatCode)}</c:formatCode>`
    + `<c:ptCount val="${values.length}"/>`
    + values.map((value, index) => `<c:pt idx="${index}"><c:v>${value}</c:v></c:pt>`).join('')
    + '</c:numCache>';
}

function rectangularCategoryFormula(formulas: readonly string[]): string {
  const first = formulaCoordinates(formulas[0]!);
  const last = formulaCoordinates(formulas.at(-1)!);
  if (first.startRow !== last.startRow || first.endRow !== last.endRow) {
    throw new Error('Multi-level category formula rows are inconsistent');
  }
  return `Sheet1!$${first.startColumn}$${first.startRow}:$${last.endColumn}$${last.endRow}`;
}

function formulaCoordinates(value: string): {
  startColumn: string;
  startRow: string;
  endColumn: string;
  endRow: string;
} {
  const match = /^Sheet1!\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+)$/.exec(value);
  if (!match) throw new Error('Category formula is not a canonical Sheet1 range');
  return {
    startColumn: match[1]!,
    startRow: match[2]!,
    endColumn: match[3]!,
    endRow: match[4]!,
  };
}

function renderAxes(
  type: ChartType,
  ids: readonly number[],
  secondary: boolean,
  categoryOptions: Readonly<ChartCategoryAxisOptions> | undefined,
  valueOptions: Readonly<ChartValueAxisOptions> | undefined,
  seriesOptions: Readonly<ChartSeriesAxisOptions> | undefined,
): string {
  if (ids.length === 0) return '';
  if (type === 'scatter' || type === 'bubble') {
    return renderHorizontalValueAxis(ids[0]!, ids[1]!, secondary, categoryOptions)
      + renderValueAxis(ids[1]!, ids[0]!, 'midCat', secondary, valueOptions);
  }
  const category = renderCategoryAxis(ids[0]!, ids[1]!, secondary, categoryOptions);
  const value = renderValueAxis(ids[1]!, ids[0]!, 'between', secondary, valueOptions);
  if (type !== 'bar3D') return category + value;
  return category + value + renderSeriesAxis(ids[2]!, ids[1]!, secondary, seriesOptions);
}

function renderCategoryAxis(
  id: number,
  crossId: number,
  secondary: boolean,
  options: Readonly<ChartCategoryAxisOptions> | undefined,
): string {
  const position = axisPosition(options?.position, secondary ? 'top' : 'bottom');
  const element = options?.kind === 'date' ? 'dateAx' : 'catAx';
  const tail = element === 'dateAx'
    ? '<c:auto val="1"/><c:lblOffset val="100"/>'
      + renderTimeUnit('baseTimeUnit', options?.baseTimeUnit)
      + (options?.majorUnit === undefined ? '' : `<c:majorUnit val="${options.majorUnit}"/>`)
      + renderTimeUnit('majorTimeUnit', options?.majorTimeUnit)
      + (options?.minorUnit === undefined ? '' : `<c:minorUnit val="${options.minorUnit}"/>`)
      + renderTimeUnit('minorTimeUnit', options?.minorTimeUnit)
    : '<c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/>'
      + (options?.majorUnit === undefined ? '' : `<c:majorUnit val="${options.majorUnit}"/>`)
      + (options?.minorUnit === undefined ? '' : `<c:minorUnit val="${options.minorUnit}"/>`)
      + (options?.labelFrequency === undefined
        ? ''
        : `<c:tickLblSkip val="${options.labelFrequency}"/>`);
  return `<c:${element}><c:axId val="${id}"/>${renderScaling(options)}`
    + `<c:delete val="${options?.visible === false ? 1 : 0}"/><c:axPos val="${position}"/>`
    + renderGridLine('majorGridlines', options?.majorGridLine, false)
    + renderGridLine('minorGridlines', options?.minorGridLine, false)
    + renderAxisTitle(options?.title)
    + renderNumberFormat(options?.numberFormat, true)
    + `<c:majorTickMark val="${tickMark(options?.majorTickMark, 'out')}"/>`
    + `<c:minorTickMark val="${tickMark(options?.minorTickMark, 'none')}"/>`
    + `<c:tickLblPos val="${labelPosition(options?.labelPosition)}"/>`
    + renderAxisShapeProperties(options?.line)
    + renderTextProperties(options, options?.labelRotation)
    + `<c:crossAx val="${crossId}"/>${renderCrossing(options?.crossesAt)}`
    + tail
    + (element === 'dateAx' && options?.labelFrequency !== undefined
      ? `<c:tickLblSkip val="${options.labelFrequency}"/>`
      : '')
    + (options?.multiLevelLabels === undefined
      ? ''
      : `<c:noMultiLvlLbl val="${options.multiLevelLabels ? 0 : 1}"/>`)
    + `</c:${element}>`;
}

function renderValueAxis(
  id: number,
  crossId: number,
  crossBetween: 'between' | 'midCat' = 'between',
  secondary = false,
  options: Readonly<ChartValueAxisOptions> | undefined = undefined,
): string {
  const position = axisPosition(options?.position, secondary ? 'right' : 'left');
  return `<c:valAx><c:axId val="${id}"/>${renderScaling(options)}`
    + `<c:delete val="${options?.visible === false ? 1 : 0}"/><c:axPos val="${position}"/>`
    + renderGridLine('majorGridlines', options?.majorGridLine, true)
    + renderGridLine('minorGridlines', options?.minorGridLine, false)
    + renderAxisTitle(options?.title)
    + renderNumberFormat(options?.numberFormat, false)
    + `<c:majorTickMark val="${tickMark(options?.majorTickMark, 'out')}"/>`
    + `<c:minorTickMark val="${tickMark(options?.minorTickMark, 'none')}"/>`
    + `<c:tickLblPos val="${labelPosition(options?.labelPosition)}"/>`
    + renderAxisShapeProperties(options?.line)
    + renderTextProperties(options, options?.labelRotation)
    + `<c:crossAx val="${crossId}"/>${renderCrossing(options?.crossesAt)}`
    + `<c:crossBetween val="${crossBetween}"/>`
    + (options?.majorUnit === undefined ? '' : `<c:majorUnit val="${options.majorUnit}"/>`)
    + (options?.minorUnit === undefined ? '' : `<c:minorUnit val="${options.minorUnit}"/>`)
    + renderDisplayUnits(options)
    + '</c:valAx>';
}

function renderHorizontalValueAxis(
  id: number,
  crossId: number,
  secondary: boolean,
  options: Readonly<ChartCategoryAxisOptions> | undefined,
): string {
  const position = axisPosition(options?.position, secondary ? 'top' : 'bottom');
  return `<c:valAx><c:axId val="${id}"/>${renderScaling(options)}`
    + `<c:delete val="${options?.visible === false ? 1 : 0}"/><c:axPos val="${position}"/>`
    + renderGridLine('majorGridlines', options?.majorGridLine, false)
    + renderGridLine('minorGridlines', options?.minorGridLine, false)
    + renderAxisTitle(options?.title)
    + renderNumberFormat(options?.numberFormat, false)
    + `<c:majorTickMark val="${tickMark(options?.majorTickMark, 'none')}"/>`
    + `<c:minorTickMark val="${tickMark(options?.minorTickMark, 'none')}"/>`
    + `<c:tickLblPos val="${labelPosition(options?.labelPosition)}"/>`
    + renderAxisShapeProperties(options?.line)
    + renderTextProperties(options, options?.labelRotation)
    + `<c:crossAx val="${crossId}"/>${renderCrossing(options?.crossesAt)}`
    + '<c:crossBetween val="midCat"/>'
    + (options?.majorUnit === undefined ? '' : `<c:majorUnit val="${options.majorUnit}"/>`)
    + (options?.minorUnit === undefined ? '' : `<c:minorUnit val="${options.minorUnit}"/>`)
    + '</c:valAx>';
}

function renderSeriesAxis(
  id: number,
  crossId: number,
  secondary: boolean,
  options: Readonly<ChartSeriesAxisOptions> | undefined,
): string {
  const position = axisPosition(options?.position, secondary ? 'top' : 'bottom');
  return `<c:serAx><c:axId val="${id}"/>${renderScaling(options)}`
    + `<c:delete val="${options?.visible === false ? 1 : 0}"/><c:axPos val="${position}"/>`
    + renderGridLine('majorGridlines', options?.majorGridLine, false)
    + renderGridLine('minorGridlines', options?.minorGridLine, false)
    + renderAxisTitle(options?.title)
    + renderNumberFormat(options?.numberFormat, false)
    + `<c:majorTickMark val="${tickMark(options?.majorTickMark, 'none')}"/>`
    + `<c:minorTickMark val="${tickMark(options?.minorTickMark, 'none')}"/>`
    + `<c:tickLblPos val="${labelPosition(options?.labelPosition, 'none')}"/>`
    + renderAxisShapeProperties(options?.line)
    + renderTextProperties(options, options?.labelRotation)
    + `<c:crossAx val="${crossId}"/><c:crosses val="autoZero"/>`
    + (options?.majorUnit === undefined ? '' : `<c:majorUnit val="${options.majorUnit}"/>`)
    + (options?.minorUnit === undefined ? '' : `<c:minorUnit val="${options.minorUnit}"/>`)
    + (options?.labelFrequency === undefined
      ? ''
      : `<c:tickLblSkip val="${options.labelFrequency}"/>`)
    + '</c:serAx>';
}

function renderChartTitle(
  options: Readonly<ChartTitleOptions> | undefined,
  language: string | undefined,
): string {
  if (!options || options.visible === false) return '';
  const text = options.text ?? 'Chart Title';
  const layout = options.position === undefined
    ? '<c:layout/>'
    : '<c:layout><c:manualLayout><c:xMode val="edge"/><c:yMode val="edge"/>'
      + `<c:x val="${options.position.x}"/><c:y val="${options.position.y}"/>`
      + '</c:manualLayout></c:layout>';
  return '<c:title><c:tx><c:rich>'
    + `<a:bodyPr${options.rotation === undefined ? '' : ` rot="${Math.round(options.rotation * 60_000)}"`}/>`
    + '<a:lstStyle/><a:p><a:r>'
    + renderRunProperties(options, language)
    + `<a:t>${escapeXmlText(text)}</a:t></a:r></a:p></c:rich></c:tx>`
    + `${layout}<c:overlay val="${options.overlay ? 1 : 0}"/></c:title>`;
}

function renderLegend(
  options: Readonly<ChartLegendOptions> | undefined,
  language: string | undefined,
): string {
  if (!options || options.visible === false) return '';
  const position = {
    bottom: 'b',
    left: 'l',
    right: 'r',
    top: 't',
    topRight: 'tr',
  }[options.position ?? 'right'];
  return `<c:legend><c:legendPos val="${position}"/><c:layout/>`
    + `<c:overlay val="${options.overlay ? 1 : 0}"/>`
    + renderTextProperties(options, undefined, language)
    + '</c:legend>';
}

function renderView3D(definition: Readonly<ChartDefinition>): string {
  const options = definition.options;
  const is3D = definition.groups[0]?.type === 'bar3D';
  if (
    !is3D
    && options.rotationX === undefined
    && options.rotationY === undefined
    && options.rightAngleAxes === undefined
    && options.perspective === undefined
  ) return '';
  return '<c:view3D>'
    + `<c:rotX val="${options.rotationX ?? 15}"/>`
    + `<c:rotY val="${options.rotationY ?? 20}"/>`
    + `<c:rAngAx val="${options.rightAngleAxes === false ? 0 : 1}"/>`
    + `<c:perspective val="${options.perspective ?? 30}"/>`
    + '</c:view3D>';
}

function renderAreaShapeProperties(
  options: Readonly<ChartAreaOptions> | undefined,
  renderDefault: boolean,
): string {
  if (!options) {
    return renderDefault
      ? '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>'
      : '';
  }
  return '<c:spPr>'
    + (options.fill ? renderSimpleFill(options.fill as SimpleFill, 'a:') : '')
    + (options.line ? renderShapeLine(options.line as NormalizedSimpleLine) : '')
    + '</c:spPr>';
}

function renderDataTable(
  options: Readonly<ChartDataTableOptions> | undefined,
  language: string | undefined,
): string {
  if (!options || options.visible === false) return '';
  return '<c:dTable>'
    + `<c:showHorzBorder val="${options.showHorizontalBorder === false ? 0 : 1}"/>`
    + `<c:showVertBorder val="${options.showVerticalBorder === false ? 0 : 1}"/>`
    + `<c:showOutline val="${options.showOutline === false ? 0 : 1}"/>`
    + `<c:showKeys val="${options.showLegendKeys === false ? 0 : 1}"/>`
    + renderTextProperties(options, undefined, language)
    + '</c:dTable>';
}

function renderDataLabels(options: Readonly<ChartDataLabelOptions> | undefined): string {
  if (!options) return '';
  const position = options.position === undefined
    ? ''
    : `<c:dLblPos val="${dataLabelPosition(options.position)}"/>`;
  return '<c:dLbls>'
    + (options.numberFormat === undefined
      ? ''
      : `<c:numFmt formatCode="${escapeXmlAttribute(options.numberFormat)}" sourceLinked="0"/>`)
    + renderTextProperties(options)
    + position
    + `<c:showLegendKey val="0"/><c:showVal val="${options.showValue ? 1 : 0}"/>`
    + `<c:showCatName val="${options.showCategoryName ? 1 : 0}"/>`
    + `<c:showSerName val="${options.showSeriesName ? 1 : 0}"/>`
    + `<c:showPercent val="${options.showPercent ? 1 : 0}"/>`
    + `<c:showBubbleSize val="${options.showBubbleSize ? 1 : 0}"/>`
    + `<c:showLeaderLines val="${options.showLeaderLines ? 1 : 0}"/>`
    + '</c:dLbls>';
}

function renderSeriesShapeProperties(
  type: ChartType,
  index: number,
  options: Readonly<ChartSeriesOptions> | undefined,
  colors: readonly RichTextColor[] | undefined,
): string {
  let fill = options?.fill as SimpleFill | undefined;
  let line = options?.line as NormalizedSimpleLine | undefined;
  const color = colors?.[index % colors.length];
  if (color && !fill && !line) {
    if (type === 'line' || type === 'radar' || type === 'scatter') {
      line = { kind: 'line', color, width: 2, dash: 'solid' };
    } else {
      fill = { kind: 'solid', color };
    }
  }
  if (!fill && !line) return '';
  return '<c:spPr>'
    + (fill ? renderSimpleFill(fill, 'a:') : '')
    + (line ? renderShapeLine(line) : '')
    + '</c:spPr>';
}

function renderMarker(options: Readonly<ChartMarkerOptions> | undefined): string {
  if (!options) return '';
  const fill = options.fill as SimpleFill | undefined;
  const line = options.line as NormalizedSimpleLine | undefined;
  const properties = !fill && !line
    ? ''
    : '<c:spPr>'
      + (fill ? renderSimpleFill(fill, 'a:') : '')
      + (line ? renderShapeLine(line) : '')
      + '</c:spPr>';
  return `<c:marker><c:symbol val="${options.shape ?? 'circle'}"/>`
    + `<c:size val="${options.size ?? 5}"/>${properties}</c:marker>`;
}

function renderShapeLine(line: NormalizedSimpleLine): string {
  if (line.kind === 'none') return '<a:ln><a:noFill/></a:ln>';
  return `<a:ln w="${Math.round(line.width * EMU_PER_POINT)}">`
    + renderSimpleLine(line, 'a:')
    + '</a:ln>';
}

function renderAxisTitle(options: Readonly<ChartTitleOptions> | undefined): string {
  if (!options || options.visible === false) return '';
  return renderChartTitle(options, undefined);
}

function renderScaling(options: Readonly<ChartAxisOptions & {
  readonly logarithmicBase?: number;
  readonly maximum?: number;
  readonly minimum?: number;
}> | undefined): string {
  return '<c:scaling>'
    + (options?.logarithmicBase === undefined
      ? ''
      : `<c:logBase val="${options.logarithmicBase}"/>`)
    + `<c:orientation val="${options?.orientation ?? 'minMax'}"/>`
    + (options?.maximum === undefined ? '' : `<c:max val="${options.maximum}"/>`)
    + (options?.minimum === undefined ? '' : `<c:min val="${options.minimum}"/>`)
    + '</c:scaling>';
}

function renderCrossing(value: number | 'autoZero' | undefined): string {
  return typeof value === 'number'
    ? `<c:crossesAt val="${value}"/>`
    : '<c:crosses val="autoZero"/>';
}

function renderTimeUnit(
  name: 'baseTimeUnit' | 'majorTimeUnit' | 'minorTimeUnit',
  value: import('./chart.js').ChartTimeUnit | undefined,
): string {
  return value === undefined ? '' : `<c:${name} val="${value}"/>`;
}

function renderDisplayUnits(options: Readonly<ChartValueAxisOptions> | undefined): string {
  if (options?.displayUnit === undefined) return '';
  return `<c:dispUnits><c:builtInUnit val="${options.displayUnit}"/>`
    + (options.displayUnitLabel === true ? '<c:dispUnitsLbl><c:layout/></c:dispUnitsLbl>' : '')
    + '</c:dispUnits>';
}

function renderGridLine(
  name: 'majorGridlines' | 'minorGridlines',
  line: import('./preset-shape.js').ShapeLine | undefined,
  renderDefault: boolean,
): string {
  if (!line) return renderDefault ? `<c:${name}/>` : '';
  return `<c:${name}><c:spPr>${renderShapeLine(line as NormalizedSimpleLine)}</c:spPr></c:${name}>`;
}

function renderNumberFormat(value: string | undefined, sourceLinked: boolean): string {
  return `<c:numFmt formatCode="${escapeXmlAttribute(value ?? 'General')}" sourceLinked="${sourceLinked ? 1 : 0}"/>`;
}

function renderAxisShapeProperties(
  line: import('./preset-shape.js').ShapeLine | undefined,
): string {
  return line ? `<c:spPr>${renderShapeLine(line as NormalizedSimpleLine)}</c:spPr>` : '';
}

function renderTextProperties(
  options: Readonly<ChartFontOptions> | undefined,
  rotation: number | undefined = undefined,
  language: string | undefined = undefined,
): string {
  if (!options || (!hasFont(options) && rotation === undefined)) return '';
  return '<c:txPr>'
    + `<a:bodyPr${rotation === undefined ? '' : ` rot="${Math.round(rotation * 60_000)}"`}/>`
    + '<a:lstStyle/><a:p><a:pPr><a:defRPr'
    + renderFontAttributes(options)
    + '>'
    + renderFontChildren(options)
    + '</a:defRPr></a:pPr>'
    + `<a:endParaRPr${language ? ` lang="${escapeXmlAttribute(language)}"` : ''}/>`
    + '</a:p></c:txPr>';
}

function renderRunProperties(options: Readonly<ChartFontOptions>, language: string | undefined): string {
  return `<a:rPr${language ? ` lang="${escapeXmlAttribute(language)}"` : ''}${renderFontAttributes(options)}>`
    + renderFontChildren(options)
    + '</a:rPr>';
}

function renderFontAttributes(options: Readonly<ChartFontOptions>): string {
  return (options.size === undefined ? '' : ` sz="${Math.round(options.size * 100)}"`)
    + (options.bold === undefined ? '' : ` b="${options.bold ? 1 : 0}"`)
    + (options.italic === undefined ? '' : ` i="${options.italic ? 1 : 0}"`);
}

function renderFontChildren(options: Readonly<ChartFontOptions>): string {
  return (options.color ? renderSimpleFill({ kind: 'solid', color: options.color }, 'a:') : '')
    + (options.face ? `<a:latin typeface="${escapeXmlAttribute(options.face)}"/>` : '');
}

function hasFont(options: Readonly<ChartFontOptions>): boolean {
  return options.face !== undefined
    || options.size !== undefined
    || options.bold !== undefined
    || options.italic !== undefined
    || options.color !== undefined;
}

function axisPosition(
  value: ChartAxisOptions['position'] | undefined,
  fallback: NonNullable<ChartAxisOptions['position']>,
): string {
  return { bottom: 'b', left: 'l', right: 'r', top: 't' }[value ?? fallback];
}

function tickMark(
  value: ChartAxisOptions['majorTickMark'] | undefined,
  fallback: 'none' | 'out',
): string {
  return value === undefined
    ? fallback
    : { cross: 'cross', inside: 'in', none: 'none', outside: 'out' }[value];
}

function labelPosition(
  value: ChartAxisOptions['labelPosition'] | undefined,
  fallback: 'nextTo' | 'none' = 'nextTo',
): string {
  return value === undefined
    ? fallback
    : { high: 'high', low: 'low', nextTo: 'nextTo', none: 'none' }[value];
}

function dataLabelPosition(value: NonNullable<ChartDataLabelOptions['position']>): string {
  return {
    bestFit: 'bestFit',
    bottom: 'b',
    center: 'ctr',
    insideBase: 'inBase',
    insideEnd: 'inEnd',
    left: 'l',
    outsideEnd: 'outEnd',
    right: 'r',
    top: 't',
  }[value];
}

function groupOption<T>(
  options: Readonly<ChartGroupOptions> | undefined,
  key: string,
): T | undefined {
  return options ? (options as Record<string, unknown>)[key] as T | undefined : undefined;
}

function readOptions(value: AddChartOptions | undefined): Record<string, unknown> {
  if (value === undefined) return {};
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError('Chart options must be an object');
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError('Chart options must be an ordinary object');
  }
  const result: Record<string, unknown> = {};
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !OPTION_KEYS.has(key)) {
      throw new TypeError(`Chart options contain unsupported property ${String(key)}`);
    }
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`Chart option ${key} must be a data property`);
    }
    result[key] = descriptor.value;
  }
  return result;
}

function readXmlString(value: unknown, context: string, fallback?: string): string {
  if (value === undefined && fallback !== undefined) return fallback;
  if (typeof value !== 'string') throw new TypeError(`${context} must be a string`);
  if (containsInvalidXmlCharacter(value)) {
    throw new TypeError(`${context} contains invalid XML characters`);
  }
  return value;
}

function readCoordinate(value: unknown, fallback: number, context: string): number {
  if (value === undefined) return fallback;
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  const rounded = Math.round(value);
  if (!Number.isSafeInteger(rounded)) throw new RangeError(`${context} must round to a safe integer`);
  return Object.is(rounded, -0) ? 0 : rounded;
}

function readBoolean(value: unknown, fallback: boolean, context: string): boolean {
  if (value === undefined) return fallback;
  if (typeof value !== 'boolean') throw new TypeError(`${context} must be a boolean`);
  return value;
}

function containsInvalidXmlCharacter(value: string): boolean {
  return /[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(value);
}
