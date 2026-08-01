import { escapeXmlAttribute, escapeXmlText } from '@pptx/lossless-xml';
import type {
  AddChartOptions,
  ChartCategories,
  ChartDefinition,
  ChartSeries,
  ChartType,
} from './chart.js';
import type {
  ChartWorkbookFormula,
  ChartWorkbookPlan,
} from './chart-workbook.internal.js';
import { inches, type Emu, type OoxmlAngle } from './units.js';

const CHART_NAMESPACE = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const DRAWING_NAMESPACE = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CHART_GRAPHIC_DATA_URI = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const MAX_ROTATION = 21_600_000;
const OPTION_KEYS = new Set([
  'name',
  'altText',
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
      return group.type === 'scatter' || group.type === 'bubble'
        ? renderXySeries(series, formula, index, group.type)
        : renderCategoricalSeries(series, formula, index);
    }).join('');
    return group.type === 'scatter' || group.type === 'bubble'
      ? renderXyGroup(group.type, renderedSeries, axisIds)
      : renderCategoricalGroup(group.type, renderedSeries, axisIds);
  }).join('');
  const axes = renderAllocatedAxes(definition, allocation);
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + `<c:chartSpace xmlns:c="${CHART_NAMESPACE}" xmlns:a="${DRAWING_NAMESPACE}" `
    + `xmlns:r="${RELATIONSHIP_NAMESPACE}">`
    + '<c:date1904 val="0"/><c:roundedCorners val="0"/><c:chart><c:autoTitleDeleted val="1"/>'
    + `<c:plotArea><c:layout/>${groupXml}${axes}</c:plotArea>`
    + '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>'
    + '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>'
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
  return renderAxes(type, allocation.primary, false)
    + (allocation.secondary ? renderAxes(type, allocation.secondary, true) : '');
}

export function renderChartGraphicFrame(
  shapeId: number,
  relationshipId: string,
  options: Readonly<NormalizedAddChartOptions>,
): string {
  const transformAttributes = [
    options.rotation === 0 ? '' : ` rot="${options.rotation}"`,
    options.flipHorizontal ? ' flipH="1"' : '',
    options.flipVertical ? ' flipV="1"' : '',
  ].join('');
  const description = options.altText === undefined
    ? ''
    : ` descr="${escapeXmlAttribute(options.altText)}"`;
  return `<p:graphicFrame xmlns:p="${PRESENTATION_NAMESPACE}" xmlns:a="${DRAWING_NAMESPACE}" `
    + `xmlns:c="${CHART_NAMESPACE}" xmlns:r="${RELATIONSHIP_NAMESPACE}">`
    + '<p:nvGraphicFramePr>'
    + `<p:cNvPr id="${shapeId}" name="${escapeXmlAttribute(options.name)}"${description}/>`
    + '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/>'
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
): string {
  const axes = axisIds.map((id) => `<c:axId val="${id}"/>`).join('');
  switch (type) {
    case 'area':
      return `<c:areaChart><c:grouping val="standard"/><c:varyColors val="0"/>${series}${axes}</c:areaChart>`;
    case 'bar':
      return `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${series}<c:gapWidth val="150"/><c:overlap val="0"/>${axes}</c:barChart>`;
    case 'bar3D':
      return `<c:bar3DChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${series}<c:gapWidth val="150"/><c:gapDepth val="150"/>${axes}</c:bar3DChart>`;
    case 'doughnut':
      return `<c:doughnutChart><c:varyColors val="1"/>${series}<c:firstSliceAng val="0"/><c:holeSize val="50"/></c:doughnutChart>`;
    case 'line':
      return `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${series}${axes}</c:lineChart>`;
    case 'pie':
      return `<c:pieChart><c:varyColors val="1"/>${series}<c:firstSliceAng val="0"/></c:pieChart>`;
    case 'radar':
      return `<c:radarChart><c:radarStyle val="standard"/><c:varyColors val="0"/>${series}${axes}</c:radarChart>`;
  }
}

function renderCategoricalSeries(
  series: Readonly<ChartSeries>,
  formula: Readonly<ChartWorkbookFormula>,
  index: number,
): string {
  return '<c:ser>'
    + `<c:idx val="${index}"/><c:order val="${index}"/>`
    + renderSeriesName(series.name, formula.name)
    + renderCategories(series.categories!, formula.categories!)
    + '<c:val><c:numRef>'
    + `<c:f>${escapeXmlText(formula.values)}</c:f>${renderNumericCache(series.values)}`
    + '</c:numRef></c:val></c:ser>';
}

function renderXyGroup(
  type: 'scatter' | 'bubble',
  series: string,
  axisIds: readonly number[],
): string {
  const axes = axisIds.map((id) => `<c:axId val="${id}"/>`).join('');
  if (type === 'scatter') {
    return '<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>'
      + `${series}${axes}</c:scatterChart>`;
  }
  return '<c:bubbleChart><c:varyColors val="0"/>'
    + `${series}<c:bubbleScale val="100"/><c:showNegBubbles val="0"/>`
    + `<c:sizeRepresents val="area"/>${axes}</c:bubbleChart>`;
}

function renderXySeries(
  series: Readonly<ChartSeries>,
  formula: Readonly<ChartWorkbookFormula>,
  index: number,
  type: 'scatter' | 'bubble',
): string {
  if (!formula.xValues || (type === 'bubble' && !formula.sizes)) {
    throw new Error(`Chart ${type} series ${index} has an incomplete workbook formula plan`);
  }
  return '<c:ser>'
    + `<c:idx val="${index}"/><c:order val="${index}"/>`
    + renderSeriesName(series.name, formula.name)
    + '<c:xVal><c:numRef>'
    + `<c:f>${escapeXmlText(formula.xValues)}</c:f>${renderNumericCache(series.xValues!)}`
    + '</c:numRef></c:xVal>'
    + '<c:yVal><c:numRef>'
    + `<c:f>${escapeXmlText(formula.values)}</c:f>${renderNumericCache(series.values)}`
    + '</c:numRef></c:yVal>'
    + (type === 'bubble'
      ? '<c:bubbleSize><c:numRef>'
        + `<c:f>${escapeXmlText(formula.sizes!)}</c:f>${renderNumericCache(series.sizes!)}`
        + '</c:numRef></c:bubbleSize><c:bubble3D val="0"/>'
      : '<c:smooth val="0"/>')
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

function renderNumericCache(values: readonly number[]): string {
  return '<c:numCache><c:formatCode>General</c:formatCode>'
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

function renderAxes(type: ChartType, ids: readonly number[], secondary: boolean): string {
  if (ids.length === 0) return '';
  if (type === 'scatter' || type === 'bubble') {
    return renderHorizontalValueAxis(ids[0]!, ids[1]!, secondary)
      + renderValueAxis(ids[1]!, ids[0]!, 'midCat', secondary);
  }
  const category = renderCategoryAxis(ids[0]!, ids[1]!, secondary);
  const value = renderValueAxis(ids[1]!, ids[0]!, 'between', secondary);
  if (type !== 'bar3D') return category + value;
  return category + value + renderSeriesAxis(ids[2]!, ids[1]!);
}

function renderCategoryAxis(id: number, crossId: number, secondary: boolean): string {
  return `<c:catAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling>`
    + `<c:delete val="0"/><c:axPos val="${secondary ? 't' : 'b'}"/>`
    + '<c:numFmt formatCode="General" sourceLinked="1"/>'
    + '<c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
    + `<c:crossAx val="${crossId}"/><c:crosses val="autoZero"/><c:auto val="1"/>`
    + '<c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx>';
}

function renderValueAxis(
  id: number,
  crossId: number,
  crossBetween: 'between' | 'midCat' = 'between',
  secondary = false,
): string {
  return `<c:valAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling>`
    + `<c:delete val="0"/><c:axPos val="${secondary ? 'r' : 'l'}"/><c:majorGridlines/>`
    + '<c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="out"/>'
    + '<c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
    + `<c:crossAx val="${crossId}"/><c:crosses val="autoZero"/><c:crossBetween val="${crossBetween}"/>`
    + '</c:valAx>';
}

function renderHorizontalValueAxis(id: number, crossId: number, secondary: boolean): string {
  return `<c:valAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling>`
    + `<c:delete val="0"/><c:axPos val="${secondary ? 't' : 'b'}"/>`
    + '<c:numFmt formatCode="General" sourceLinked="0"/>'
    + '<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
    + `<c:crossAx val="${crossId}"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/>`
    + '</c:valAx>';
}

function renderSeriesAxis(id: number, crossId: number): string {
  return `<c:serAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling>`
    + '<c:delete val="0"/><c:axPos val="b"/><c:majorTickMark val="none"/>'
    + '<c:minorTickMark val="none"/><c:tickLblPos val="none"/>'
    + `<c:crossAx val="${crossId}"/><c:crosses val="autoZero"/></c:serAx>`;
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
