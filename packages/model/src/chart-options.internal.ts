import type {
  ChartAreaOptions,
  ChartAxisOptions,
  ChartCategoryAxisOptions,
  ChartDataLabelOptions,
  ChartDataTableOptions,
  ChartDisplayUnit,
  ChartFontOptions,
  ChartGroupOptions,
  ChartLegendOptions,
  ChartMarkerOptions,
  ChartOptions,
  ChartPointDataLabelOptions,
  ChartPointOptions,
  ChartSeriesAxisOptions,
  ChartSeriesDataLabelOptions,
  ChartSeriesOptions,
  ChartTimeUnit,
  ChartTitleOptions,
  ChartType,
  ChartValueAxisOptions,
} from './chart.js';
import { normalizeShapeShadow } from './simple-shadow.internal.js';
import { normalizeSimpleFill, type SimpleFill } from './simple-fill.internal.js';
import {
  normalizeSimpleLine,
  type NormalizedSimpleLine,
} from './simple-line.internal.js';
import type { RichTextColor } from './text.js';

const OPTION_KEYS = [
  'language',
  'style',
  'roundedCorners',
  'displayBlanksAs',
  'title',
  'legend',
  'chartArea',
  'plotArea',
  'layout',
  'categoryAxis',
  'valueAxis',
  'seriesAxis',
  'secondaryCategoryAxis',
  'secondaryValueAxis',
  'dataTable',
  'colors',
  'rightAngleAxes',
  'rotationX',
  'rotationY',
  'perspective',
] as const;
const FONT_KEYS = ['face', 'size', 'bold', 'italic', 'color'] as const;
const TITLE_KEYS = [
  ...FONT_KEYS,
  'visible',
  'text',
  'overlay',
  'rotation',
  'position',
  'align',
] as const;
const TITLE_POSITION_KEYS = ['x', 'y'] as const;
const LAYOUT_KEYS = ['x', 'y', 'width', 'height'] as const;
const LEGEND_KEYS = [...FONT_KEYS, 'visible', 'position', 'overlay'] as const;
const AREA_KEYS = ['fill', 'line'] as const;
const AXIS_KEYS = [
  ...FONT_KEYS,
  'visible',
  'position',
  'title',
  'numberFormat',
  'orientation',
  'labelPosition',
  'labelRotation',
  'line',
  'majorGridLine',
  'minorGridLine',
  'majorTickMark',
  'minorTickMark',
] as const;
const CATEGORY_AXIS_KEYS = [
  ...AXIS_KEYS,
  'kind',
  'minimum',
  'maximum',
  'majorUnit',
  'minorUnit',
  'crossesAt',
  'baseTimeUnit',
  'majorTimeUnit',
  'minorTimeUnit',
  'labelFrequency',
  'multiLevelLabels',
] as const;
const VALUE_AXIS_KEYS = [
  ...AXIS_KEYS,
  'minimum',
  'maximum',
  'majorUnit',
  'minorUnit',
  'logarithmicBase',
  'crossesAt',
  'displayUnit',
  'displayUnitLabel',
] as const;
const SERIES_AXIS_KEYS = [
  ...AXIS_KEYS,
  'majorUnit',
  'minorUnit',
  'labelFrequency',
] as const;
const TIME_UNITS = ['days', 'months', 'years'] as const satisfies readonly ChartTimeUnit[];
const DISPLAY_UNITS = [
  'billions',
  'hundredMillions',
  'hundredThousands',
  'hundreds',
  'millions',
  'tenMillions',
  'tenThousands',
  'thousands',
  'trillions',
] as const satisfies readonly ChartDisplayUnit[];
const DATA_LABEL_KEYS = [
  ...FONT_KEYS,
  'showValue',
  'showCategoryName',
  'showSeriesName',
  'showPercent',
  'showBubbleSize',
  'position',
  'numberFormat',
  'showLeaderLines',
] as const;
const DATA_TABLE_KEYS = [
  ...FONT_KEYS,
  'visible',
  'showHorizontalBorder',
  'showVerticalBorder',
  'showOutline',
  'showLegendKeys',
  'numberFormat',
] as const;
const MARKER_KEYS = ['shape', 'size', 'fill', 'line'] as const;
const POINT_KEYS = ['index', 'fill', 'line', 'shadow'] as const;
const POINT_DATA_LABEL_KEYS = ['index', 'text', 'fields'] as const;
const SERIES_DATA_LABEL_KEYS = [...DATA_LABEL_KEYS, 'fill', 'pointLabels'] as const;
const SERIES_KEYS = ['fill', 'line', 'marker', 'shadow', 'points', 'dataLabels'] as const;
const COMMON_GROUP_KEYS = ['varyColors', 'series', 'dataLabels'] as const;

const GROUP_KEYS: Readonly<Record<ChartType, readonly string[]>> = Object.freeze({
  area: [...COMMON_GROUP_KEYS, 'grouping'],
  bar: [...COMMON_GROUP_KEYS, 'direction', 'grouping', 'gapWidth', 'overlap'],
  bar3D: [
    ...COMMON_GROUP_KEYS,
    'direction',
    'grouping',
    'gapWidth',
    'gapDepth',
    'shape',
  ],
  bubble: [...COMMON_GROUP_KEYS, 'scale', 'showNegativeBubbles', 'sizeRepresents'],
  bubble3D: [...COMMON_GROUP_KEYS, 'scale', 'showNegativeBubbles', 'sizeRepresents'],
  doughnut: [...COMMON_GROUP_KEYS, 'firstSliceAngle', 'holeSize'],
  line: [...COMMON_GROUP_KEYS, 'grouping', 'smooth', 'marker'],
  pie: [...COMMON_GROUP_KEYS, 'firstSliceAngle'],
  radar: [...COMMON_GROUP_KEYS, 'style', 'marker'],
  scatter: [...COMMON_GROUP_KEYS, 'style', 'smooth', 'marker'],
});

export function normalizeChartOptions(value: unknown): Readonly<ChartOptions> {
  if (value === undefined) return EMPTY_OPTIONS;
  const input = readObject(value, OPTION_KEYS, 'Chart options');
  const result: ChartOptions = {
    ...optionalString(input, 'language', 'Chart language', false),
    ...optionalInteger(input, 'style', 'Chart style', 1, 48),
    ...optionalBoolean(input, 'roundedCorners', 'Chart roundedCorners'),
    ...optionalEnum(input, 'displayBlanksAs', 'Chart displayBlanksAs', ['gap', 'span', 'zero']),
    ...optionalObject(input, 'title', normalizeTitleOptions, 'Chart title'),
    ...optionalObject(input, 'legend', normalizeLegendOptions, 'Chart legend'),
    ...optionalObject(input, 'chartArea', normalizeAreaOptions, 'Chart area'),
    ...optionalObject(input, 'plotArea', normalizeAreaOptions, 'Chart plot area'),
    ...optionalObject(input, 'layout', normalizeLayoutOptions, 'Chart layout'),
    ...optionalObject(
      input,
      'categoryAxis',
      normalizeCategoryAxisOptions,
      'Chart category axis',
    ),
    ...optionalObject(input, 'valueAxis', normalizeValueAxisOptions, 'Chart value axis'),
    ...optionalObject(input, 'seriesAxis', normalizeSeriesAxisOptions, 'Chart series axis'),
    ...optionalObject(
      input,
      'secondaryCategoryAxis',
      normalizeCategoryAxisOptions,
      'Chart secondary category axis',
    ),
    ...optionalObject(
      input,
      'secondaryValueAxis',
      normalizeValueAxisOptions,
      'Chart secondary value axis',
    ),
    ...optionalObject(input, 'dataTable', normalizeDataTableOptions, 'Chart data table'),
    ...optionalColors(input),
    ...optionalBoolean(input, 'rightAngleAxes', 'Chart rightAngleAxes'),
    ...optionalNumber(input, 'rotationX', 'Chart rotationX', -90, 90),
    ...optionalNumber(input, 'rotationY', 'Chart rotationY', 0, 360),
    ...optionalNumber(input, 'perspective', 'Chart perspective', 0, 240),
  };
  return Object.freeze(result);
}

export function normalizeChartGroupOptions(
  type: ChartType,
  value: unknown,
  seriesValueCounts: readonly number[],
): Readonly<ChartGroupOptions> | undefined {
  if (value === undefined) return undefined;
  const context = `Chart ${type} group options`;
  const input = readObject(value, GROUP_KEYS[type], context);
  const common: ChartGroupOptions = {
    ...optionalBoolean(input, 'varyColors', `${context} varyColors`),
    ...optionalSeriesOptions(input, type, seriesValueCounts, context),
    ...optionalObject(input, 'dataLabels', normalizeDataLabelOptions, `${context} dataLabels`),
  };
  let specific: object;
  switch (type) {
    case 'area':
      specific = optionalEnum(
        input,
        'grouping',
        `${context} grouping`,
        ['percentStacked', 'stacked', 'standard'],
      );
      break;
    case 'bar':
    case 'bar3D':
      specific = {
        ...optionalEnum(input, 'direction', `${context} direction`, ['bar', 'column']),
        ...optionalEnum(
          input,
          'grouping',
          `${context} grouping`,
          type === 'bar3D'
            ? ['clustered', 'percentStacked', 'stacked', 'standard']
            : ['clustered', 'percentStacked', 'stacked'],
        ),
        ...optionalInteger(input, 'gapWidth', `${context} gapWidth`, 0, 500),
        ...(type === 'bar'
          ? optionalInteger(input, 'overlap', `${context} overlap`, -100, 100)
          : {}),
        ...(type === 'bar3D'
          ? {
              ...optionalInteger(input, 'gapDepth', `${context} gapDepth`, 0, 500),
              ...optionalEnum(input, 'shape', `${context} shape`, [
                'box',
                'cone',
                'coneToMax',
                'cylinder',
                'pyramid',
                'pyramidToMax',
              ]),
            }
          : {}),
      };
      break;
    case 'bubble':
    case 'bubble3D':
      specific = {
        ...optionalInteger(input, 'scale', `${context} scale`, 0, 300),
        ...optionalBoolean(input, 'showNegativeBubbles', `${context} showNegativeBubbles`),
        ...optionalEnum(input, 'sizeRepresents', `${context} sizeRepresents`, ['area', 'width']),
      };
      break;
    case 'doughnut':
      specific = {
        ...optionalInteger(input, 'firstSliceAngle', `${context} firstSliceAngle`, 0, 360),
        ...optionalInteger(input, 'holeSize', `${context} holeSize`, 10, 90),
      };
      break;
    case 'line':
      specific = {
        ...optionalEnum(
          input,
          'grouping',
          `${context} grouping`,
          ['percentStacked', 'stacked', 'standard'],
        ),
        ...optionalBoolean(input, 'smooth', `${context} smooth`),
        ...optionalObject(input, 'marker', normalizeMarkerOptions, `${context} marker`),
      };
      break;
    case 'pie':
      specific = optionalInteger(
        input,
        'firstSliceAngle',
        `${context} firstSliceAngle`,
        0,
        360,
      );
      break;
    case 'radar':
      specific = {
        ...optionalEnum(input, 'style', `${context} style`, ['filled', 'marker', 'standard']),
        ...optionalObject(input, 'marker', normalizeMarkerOptions, `${context} marker`),
      };
      break;
    case 'scatter':
      specific = {
        ...optionalEnum(
          input,
          'style',
          `${context} style`,
          ['line', 'lineMarker', 'marker', 'none', 'smooth', 'smoothMarker'],
        ),
        ...optionalBoolean(input, 'smooth', `${context} smooth`),
        ...optionalObject(input, 'marker', normalizeMarkerOptions, `${context} marker`),
      };
      break;
  }
  const result = Object.freeze({ ...common, ...specific }) as Readonly<ChartGroupOptions>;
  validateDataLabelCompatibility(type, result.dataLabels);
  return result;
}

export function chartOptionValuesEqual(left: unknown, right: unknown): boolean {
  if (left === right) return true;
  if (Array.isArray(left) || Array.isArray(right)) {
    return Array.isArray(left)
      && Array.isArray(right)
      && left.length === right.length
      && left.every((value, index) => chartOptionValuesEqual(value, right[index]));
  }
  if (!left || !right || typeof left !== 'object' || typeof right !== 'object') return false;
  const leftKeys = Object.keys(left);
  const rightKeys = Object.keys(right);
  return leftKeys.length === rightKeys.length
    && leftKeys.every((key, index) =>
      key === rightKeys[index]
      && chartOptionValuesEqual(
        (left as Record<string, unknown>)[key],
        (right as Record<string, unknown>)[key],
      ));
}

const EMPTY_OPTIONS: Readonly<ChartOptions> = Object.freeze({});

function normalizeTitleOptions(value: unknown, context: string): Readonly<ChartTitleOptions> {
  const input = readObject(value, TITLE_KEYS, context);
  return Object.freeze({
    ...normalizeFontOptions(input, context),
    ...optionalBoolean(input, 'visible', `${context} visible`),
    ...optionalString(input, 'text', `${context} text`, true),
    ...optionalBoolean(input, 'overlay', `${context} overlay`),
    ...optionalNumber(input, 'rotation', `${context} rotation`, -90, 90),
    ...optionalObject(input, 'position', normalizeTitlePosition, `${context} position`),
    ...optionalEnum(input, 'align', `${context} align`, ['left', 'center', 'right']),
  });
}

function normalizeLayoutOptions(
  value: unknown,
  context: string,
): Readonly<{ readonly x: number; readonly y: number; readonly width: number; readonly height: number }> {
  const input = readObject(value, LAYOUT_KEYS, context);
  for (const key of LAYOUT_KEYS) {
    if (!Object.hasOwn(input, key)) throw new TypeError(`${context} must provide ${key}`);
  }
  return Object.freeze({
    ...optionalNumber(input, 'x', `${context} x`, 0, 1),
    ...optionalNumber(input, 'y', `${context} y`, 0, 1),
    ...optionalNumber(input, 'width', `${context} width`, 0, 1),
    ...optionalNumber(input, 'height', `${context} height`, 0, 1),
  }) as Readonly<{
    readonly x: number;
    readonly y: number;
    readonly width: number;
    readonly height: number;
  }>;
}

function normalizeTitlePosition(
  value: unknown,
  context: string,
): Readonly<{ readonly x: number; readonly y: number }> {
  const input = readObject(value, TITLE_POSITION_KEYS, context);
  if (!Object.hasOwn(input, 'x') || !Object.hasOwn(input, 'y')) {
    throw new TypeError(`${context} must provide x and y`);
  }
  return Object.freeze({
    ...optionalNumber(input, 'x', `${context} x`, 0, 1),
    ...optionalNumber(input, 'y', `${context} y`, 0, 1),
  }) as Readonly<{ readonly x: number; readonly y: number }>;
}

function normalizeLegendOptions(value: unknown, context: string): Readonly<ChartLegendOptions> {
  const input = readObject(value, LEGEND_KEYS, context);
  return Object.freeze({
    ...normalizeFontOptions(input, context),
    ...optionalBoolean(input, 'visible', `${context} visible`),
    ...optionalEnum(
      input,
      'position',
      `${context} position`,
      ['bottom', 'left', 'right', 'top', 'topRight'],
    ),
    ...optionalBoolean(input, 'overlay', `${context} overlay`),
  });
}

function normalizeAreaOptions(value: unknown, context: string): Readonly<ChartAreaOptions> {
  const input = readObject(value, AREA_KEYS, context);
  const fill = normalizeSimpleFill(input.fill, `${context} fill`);
  const line = normalizeSimpleLine(input.line, `${context} line`);
  return Object.freeze({
    ...(fill === undefined ? {} : { fill: freezeFill(fill) }),
    ...(line === undefined ? {} : { line: freezeLine(line) }),
  });
}

function normalizeCategoryAxisOptions(
  value: unknown,
  context: string,
): Readonly<ChartCategoryAxisOptions> {
  const input = readObject(value, CATEGORY_AXIS_KEYS, context);
  const kind = optionalEnum(input, 'kind', `${context} kind`, ['category', 'date']);
  const timeUnits = {
    ...optionalEnum(input, 'baseTimeUnit', `${context} baseTimeUnit`, TIME_UNITS),
    ...optionalEnum(input, 'majorTimeUnit', `${context} majorTimeUnit`, TIME_UNITS),
    ...optionalEnum(input, 'minorTimeUnit', `${context} minorTimeUnit`, TIME_UNITS),
  };
  const hasTimeUnits = Object.keys(timeUnits).length > 0;
  if (input.kind === 'category' && hasTimeUnits) {
    throw new TypeError(`${context} category kind cannot use time units`);
  }
  const result: ChartCategoryAxisOptions = {
    ...normalizeAxisBaseOptions(input, context),
    ...kind,
    ...(!Object.hasOwn(kind, 'kind') && hasTimeUnits ? { kind: 'date' as const } : {}),
    ...optionalFinite(input, 'minimum', `${context} minimum`),
    ...optionalFinite(input, 'maximum', `${context} maximum`),
    ...optionalPositive(input, 'majorUnit', `${context} majorUnit`),
    ...optionalPositive(input, 'minorUnit', `${context} minorUnit`),
    ...optionalCrossesAt(input, context),
    ...timeUnits,
    ...optionalInteger(
      input,
      'labelFrequency',
      `${context} labelFrequency`,
      1,
      Number.MAX_SAFE_INTEGER,
    ),
    ...optionalBoolean(input, 'multiLevelLabels', `${context} multiLevelLabels`),
  };
  validateAxisRange(result, context);
  return Object.freeze(result);
}

function normalizeValueAxisOptions(
  value: unknown,
  context: string,
): Readonly<ChartValueAxisOptions> {
  const input = readObject(value, VALUE_AXIS_KEYS, context);
  const result: ChartValueAxisOptions = {
    ...normalizeAxisBaseOptions(input, context),
    ...optionalFinite(input, 'minimum', `${context} minimum`),
    ...optionalFinite(input, 'maximum', `${context} maximum`),
    ...optionalPositive(input, 'majorUnit', `${context} majorUnit`),
    ...optionalPositive(input, 'minorUnit', `${context} minorUnit`),
    ...optionalNumber(input, 'logarithmicBase', `${context} logarithmicBase`, 2, 1_000),
    ...optionalCrossesAt(input, context),
    ...optionalEnum(input, 'displayUnit', `${context} displayUnit`, DISPLAY_UNITS),
    ...optionalBoolean(input, 'displayUnitLabel', `${context} displayUnitLabel`),
  };
  if (result.displayUnitLabel !== undefined && result.displayUnit === undefined) {
    throw new TypeError(`${context} displayUnitLabel requires displayUnit`);
  }
  validateAxisRange(result, context);
  return Object.freeze(result);
}

function normalizeSeriesAxisOptions(
  value: unknown,
  context: string,
): Readonly<ChartSeriesAxisOptions> {
  const input = readObject(value, SERIES_AXIS_KEYS, context);
  return Object.freeze({
    ...normalizeAxisBaseOptions(input, context),
    ...optionalPositive(input, 'majorUnit', `${context} majorUnit`),
    ...optionalPositive(input, 'minorUnit', `${context} minorUnit`),
    ...optionalInteger(
      input,
      'labelFrequency',
      `${context} labelFrequency`,
      1,
      Number.MAX_SAFE_INTEGER,
    ),
  });
}

function normalizeAxisBaseOptions(
  input: Record<string, unknown>,
  context: string,
): Readonly<ChartAxisOptions> {
  return Object.freeze({
    ...normalizeFontOptions(input, context),
    ...optionalBoolean(input, 'visible', `${context} visible`),
    ...optionalEnum(input, 'position', `${context} position`, ['bottom', 'left', 'right', 'top']),
    ...optionalObject(input, 'title', normalizeTitleOptions, `${context} title`),
    ...optionalString(input, 'numberFormat', `${context} numberFormat`, true),
    ...optionalEnum(input, 'orientation', `${context} orientation`, ['maxMin', 'minMax']),
    ...optionalEnum(
      input,
      'labelPosition',
      `${context} labelPosition`,
      ['high', 'low', 'nextTo', 'none'],
    ),
    ...optionalNumber(input, 'labelRotation', `${context} labelRotation`, -90, 90),
    ...optionalLine(input, 'line', `${context} line`),
    ...optionalLine(input, 'majorGridLine', `${context} majorGridLine`),
    ...optionalLine(input, 'minorGridLine', `${context} minorGridLine`),
    ...optionalEnum(
      input,
      'majorTickMark',
      `${context} majorTickMark`,
      ['cross', 'inside', 'none', 'outside'],
    ),
    ...optionalEnum(
      input,
      'minorTickMark',
      `${context} minorTickMark`,
      ['cross', 'inside', 'none', 'outside'],
    ),
  });
}

function validateAxisRange(
  options: Readonly<{ readonly minimum?: number; readonly maximum?: number }>,
  context: string,
): void {
  if (
    options.minimum !== undefined
    && options.maximum !== undefined
    && options.minimum >= options.maximum
  ) throw new RangeError(`${context} minimum must be less than maximum`);
}

function optionalCrossesAt(input: Record<string, unknown>, context: string): object {
  if (!Object.hasOwn(input, 'crossesAt')) return {};
  if (input.crossesAt === 'autoZero') return { crossesAt: 'autoZero' as const };
  if (typeof input.crossesAt !== 'number' || !Number.isFinite(input.crossesAt)) {
    throw new TypeError(`${context} crossesAt must be a finite number or autoZero`);
  }
  return { crossesAt: input.crossesAt };
}

function normalizeDataLabelOptions(
  value: unknown,
  context: string,
): Readonly<ChartDataLabelOptions> {
  const input = readObject(value, DATA_LABEL_KEYS, context);
  return normalizeDataLabelRecord(input, context);
}

function normalizeDataLabelRecord(
  input: Record<string, unknown>,
  context: string,
): Readonly<ChartDataLabelOptions> {
  return Object.freeze({
    ...normalizeFontOptions(input, context),
    ...optionalBoolean(input, 'showValue', `${context} showValue`),
    ...optionalBoolean(input, 'showCategoryName', `${context} showCategoryName`),
    ...optionalBoolean(input, 'showSeriesName', `${context} showSeriesName`),
    ...optionalBoolean(input, 'showPercent', `${context} showPercent`),
    ...optionalBoolean(input, 'showBubbleSize', `${context} showBubbleSize`),
    ...optionalEnum(
      input,
      'position',
      `${context} position`,
      [
        'bestFit',
        'bottom',
        'center',
        'insideBase',
        'insideEnd',
        'left',
        'outsideEnd',
        'right',
        'top',
      ],
    ),
    ...optionalString(input, 'numberFormat', `${context} numberFormat`, true),
    ...optionalBoolean(input, 'showLeaderLines', `${context} showLeaderLines`),
  });
}

function normalizeDataTableOptions(
  value: unknown,
  context: string,
): Readonly<ChartDataTableOptions> {
  const input = readObject(value, DATA_TABLE_KEYS, context);
  return Object.freeze({
    ...normalizeFontOptions(input, context),
    ...optionalBoolean(input, 'visible', `${context} visible`),
    ...optionalBoolean(input, 'showHorizontalBorder', `${context} showHorizontalBorder`),
    ...optionalBoolean(input, 'showVerticalBorder', `${context} showVerticalBorder`),
    ...optionalBoolean(input, 'showOutline', `${context} showOutline`),
    ...optionalBoolean(input, 'showLegendKeys', `${context} showLegendKeys`),
    ...optionalString(input, 'numberFormat', `${context} numberFormat`, true),
  });
}

function normalizeMarkerOptions(value: unknown, context: string): Readonly<ChartMarkerOptions> {
  const input = readObject(value, MARKER_KEYS, context);
  const fill = normalizeSimpleFill(input.fill, `${context} fill`);
  const line = normalizeSimpleLine(input.line, `${context} line`);
  return Object.freeze({
    ...optionalEnum(
      input,
      'shape',
      `${context} shape`,
      ['circle', 'dash', 'diamond', 'dot', 'none', 'plus', 'square', 'star', 'triangle', 'x'],
    ),
    ...optionalInteger(input, 'size', `${context} size`, 2, 72),
    ...(fill === undefined ? {} : { fill: freezeFill(fill) }),
    ...(line === undefined ? {} : { line: freezeLine(line) }),
  });
}

function normalizeSeriesOptions(
  value: unknown,
  type: ChartType,
  valueCount: number,
  context: string,
): Readonly<ChartSeriesOptions> {
  const input = readObject(value, SERIES_KEYS, context);
  const fill = normalizeSimpleFill(input.fill, `${context} fill`);
  const line = normalizeSimpleLine(input.line, `${context} line`);
  const shadow = input.shadow !== undefined
    ? normalizeShapeShadow(input.shadow, `${context} shadow`)
    : undefined;
  const points = Object.hasOwn(input, 'points')
    ? normalizePointOptions(input.points, valueCount, `${context} points`)
    : undefined;
  const dataLabels = Object.hasOwn(input, 'dataLabels')
    ? normalizeSeriesDataLabelOptions(
        input.dataLabels,
        type,
        valueCount,
        `${context} dataLabels`,
      )
    : undefined;
  validateDataLabelCompatibility(type, dataLabels);
  return Object.freeze({
    ...(fill === undefined ? {} : { fill: freezeFill(fill) }),
    ...(line === undefined ? {} : { line: freezeLine(line) }),
    ...optionalObject(input, 'marker', normalizeMarkerOptions, `${context} marker`),
    ...(shadow === undefined ? {} : { shadow }),
    ...(points === undefined ? {} : { points }),
    ...(dataLabels === undefined ? {} : { dataLabels }),
  });
}

function normalizePointOptions(
  value: unknown,
  valueCount: number,
  context: string,
): readonly Readonly<ChartPointOptions>[] {
  const values = readArray(value, context);
  const indexes = new Set<number>();
  const normalized = values.map((entry, position) => {
    const itemContext = `${context} item ${position}`;
    const input = readObject(entry, POINT_KEYS, itemContext);
    if (!Object.hasOwn(input, 'index')) throw new TypeError(`${itemContext} must provide index`);
    if (!Number.isSafeInteger(input.index) || (input.index as number) < 0) {
      throw new TypeError(`${itemContext} index must be a non-negative safe integer`);
    }
    const index = input.index as number;
    if (index >= valueCount) throw new RangeError(`${itemContext} index is outside the series values`);
    if (indexes.has(index)) throw new RangeError(`${context} indexes must be unique`);
    indexes.add(index);
    const fill = normalizeSimpleFill(input.fill, `${itemContext} fill`);
    const line = normalizeSimpleLine(input.line, `${itemContext} line`);
    const shadow = input.shadow !== undefined
      ? normalizeShapeShadow(input.shadow, `${itemContext} shadow`)
      : undefined;
    return Object.freeze({
      index,
      ...(fill === undefined ? {} : { fill: freezeFill(fill) }),
      ...(line === undefined ? {} : { line: freezeLine(line) }),
      ...(shadow === undefined ? {} : { shadow }),
    });
  });
  return Object.freeze(normalized);
}

function normalizeSeriesDataLabelOptions(
  value: unknown,
  type: ChartType,
  valueCount: number,
  context: string,
): Readonly<ChartSeriesDataLabelOptions> {
  const input = readObject(value, SERIES_DATA_LABEL_KEYS, context);
  const fill = normalizeSimpleFill(input.fill, `${context} fill`);
  const pointLabels = Object.hasOwn(input, 'pointLabels')
    ? normalizePointDataLabelOptions(input.pointLabels, valueCount, `${context} pointLabels`)
    : undefined;
  if (pointLabels !== undefined && type !== 'scatter') {
    throw new TypeError(`${context} pointLabels are supported only for scatter charts`);
  }
  return Object.freeze({
    ...normalizeDataLabelRecord(input, context),
    ...(fill === undefined ? {} : { fill: freezeFill(fill) }),
    ...(pointLabels === undefined ? {} : { pointLabels }),
  });
}

function normalizePointDataLabelOptions(
  value: unknown,
  valueCount: number,
  context: string,
): readonly Readonly<ChartPointDataLabelOptions>[] {
  const values = readArray(value, context);
  const indexes = new Set<number>();
  const normalized = values.map((entry, position) => {
    const itemContext = `${context} item ${position}`;
    const input = readObject(entry, POINT_DATA_LABEL_KEYS, itemContext);
    if (!Object.hasOwn(input, 'index')) throw new TypeError(`${itemContext} must provide index`);
    if (!Number.isSafeInteger(input.index) || (input.index as number) < 0) {
      throw new TypeError(`${itemContext} index must be a non-negative safe integer`);
    }
    const index = input.index as number;
    if (index >= valueCount) throw new RangeError(`${itemContext} index is outside the series values`);
    if (indexes.has(index)) throw new RangeError(`${context} indexes must be unique`);
    indexes.add(index);
    const text = optionalString(input, 'text', `${itemContext} text`, true);
    const fields = Object.hasOwn(input, 'fields')
      ? normalizeDataLabelFields(input.fields, `${itemContext} fields`)
      : undefined;
    if (!Object.hasOwn(text, 'text') && (fields === undefined || fields.length === 0)) {
      throw new TypeError(`${itemContext} must provide text or at least one field`);
    }
    return Object.freeze({
      index,
      ...text,
      ...(fields === undefined ? {} : { fields }),
    });
  });
  return Object.freeze(normalized);
}

function normalizeDataLabelFields(
  value: unknown,
  context: string,
): readonly ('xValue' | 'yValue')[] {
  const values = readArray(value, context);
  const fields = values.map((entry, index) => {
    if (entry !== 'xValue' && entry !== 'yValue') {
      throw new TypeError(`${context} item ${index} must be xValue or yValue`);
    }
    return entry;
  });
  if (new Set(fields).size !== fields.length) {
    throw new RangeError(`${context} must not contain duplicate fields`);
  }
  return Object.freeze(fields);
}

function normalizeFontOptions(
  input: Record<string, unknown>,
  context: string,
): Readonly<ChartFontOptions> {
  return Object.freeze({
    ...optionalString(input, 'face', `${context} face`, true),
    ...optionalNumber(input, 'size', `${context} size`, 1, 400),
    ...optionalBoolean(input, 'bold', `${context} bold`),
    ...optionalBoolean(input, 'italic', `${context} italic`),
    ...optionalColor(input, 'color', `${context} color`),
  });
}

function optionalSeriesOptions(
  input: Record<string, unknown>,
  type: ChartType,
  seriesValueCounts: readonly number[],
  context: string,
): object {
  if (!Object.hasOwn(input, 'series')) return {};
  const values = readArray(input.series, `${context} series`);
  if (values.length > seriesValueCounts.length) {
    throw new RangeError(`${context} series styles cannot exceed the group series count`);
  }
  const normalized = values.map((value, index) =>
    normalizeSeriesOptions(
      value,
      type,
      seriesValueCounts[index]!,
      `${context} series ${index}`,
    ));
  if (
    type !== 'line'
    && type !== 'scatter'
    && type !== 'radar'
    && normalized.some(({ marker }) => marker !== undefined)
  ) throw new TypeError(`${context} series markers are supported only for line, radar, and scatter charts`);
  return { series: Object.freeze(normalized) };
}

function optionalColors(input: Record<string, unknown>): object {
  if (!Object.hasOwn(input, 'colors')) return {};
  const values = readArray(input.colors, 'Chart colors');
  if (values.length === 0) throw new RangeError('Chart colors must not be empty');
  return {
    colors: Object.freeze(values.map((value, index) =>
      normalizeColor(value, `Chart colors item ${index}`))),
  };
}

function optionalColor(
  input: Record<string, unknown>,
  key: string,
  context: string,
): object {
  return Object.hasOwn(input, key) ? { [key]: normalizeColor(input[key], context) } : {};
}

function optionalLine(
  input: Record<string, unknown>,
  key: string,
  context: string,
): object {
  if (!Object.hasOwn(input, key)) return {};
  const line = normalizeSimpleLine(input[key], context);
  if (!line) throw new TypeError(`${context} is required`);
  return { [key]: freezeLine(line) };
}

function optionalObject<T>(
  input: Record<string, unknown>,
  key: string,
  normalize: (value: unknown, context: string) => T,
  context: string,
): object {
  return Object.hasOwn(input, key) ? { [key]: normalize(input[key], context) } : {};
}

function optionalString(
  input: Record<string, unknown>,
  key: string,
  context: string,
  nonEmpty: boolean,
): object {
  if (!Object.hasOwn(input, key)) return {};
  const value = input[key];
  if (typeof value !== 'string') throw new TypeError(`${context} must be a string`);
  if (nonEmpty && value.length === 0) throw new RangeError(`${context} must not be empty`);
  if (containsInvalidXmlCharacter(value)) {
    throw new TypeError(`${context} contains invalid XML characters`);
  }
  return { [key]: value };
}

function optionalBoolean(
  input: Record<string, unknown>,
  key: string,
  context: string,
): object {
  if (!Object.hasOwn(input, key)) return {};
  if (typeof input[key] !== 'boolean') throw new TypeError(`${context} must be a boolean`);
  return { [key]: input[key] };
}

function optionalEnum<T extends string>(
  input: Record<string, unknown>,
  key: string,
  context: string,
  values: readonly T[],
): object {
  if (!Object.hasOwn(input, key)) return {};
  if (typeof input[key] !== 'string' || !values.includes(input[key] as T)) {
    throw new TypeError(`${context} is unsupported`);
  }
  return { [key]: input[key] };
}

function optionalInteger(
  input: Record<string, unknown>,
  key: string,
  context: string,
  minimum: number,
  maximum: number,
): object {
  if (!Object.hasOwn(input, key)) return {};
  const value = input[key];
  if (typeof value !== 'number' || !Number.isSafeInteger(value)) {
    throw new TypeError(`${context} must be a safe integer`);
  }
  if (value < minimum || value > maximum) {
    throw new RangeError(`${context} must be between ${minimum} and ${maximum}`);
  }
  return { [key]: value };
}

function optionalFinite(
  input: Record<string, unknown>,
  key: string,
  context: string,
): object {
  if (!Object.hasOwn(input, key)) return {};
  const value = input[key];
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  return { [key]: Object.is(value, -0) ? 0 : value };
}

function optionalPositive(
  input: Record<string, unknown>,
  key: string,
  context: string,
): object {
  const result = optionalFinite(input, key, context) as Record<string, number>;
  if (Object.hasOwn(result, key) && result[key]! <= 0) {
    throw new RangeError(`${context} must be greater than zero`);
  }
  return result;
}

function optionalNumber(
  input: Record<string, unknown>,
  key: string,
  context: string,
  minimum: number,
  maximum: number,
): object {
  const result = optionalFinite(input, key, context) as Record<string, number>;
  if (!Object.hasOwn(result, key)) return result;
  const value = result[key]!;
  if (value < minimum || value > maximum) {
    throw new RangeError(`${context} must be between ${minimum} and ${maximum}`);
  }
  return { [key]: Math.round(value * 60_000) / 60_000 };
}

function normalizeColor(value: unknown, context: string): Readonly<RichTextColor> {
  const fill = normalizeSimpleFill({ kind: 'solid', color: value }, context);
  if (!fill || fill.kind !== 'solid') throw new TypeError(`${context} is unsupported`);
  return Object.freeze({ ...fill.color });
}

function freezeFill(fill: SimpleFill): SimpleFill {
  return fill.kind === 'none'
    ? Object.freeze({ kind: 'none' })
    : Object.freeze({
        kind: 'solid',
        color: Object.freeze({ ...fill.color }),
        ...(fill.transparency === undefined ? {} : { transparency: fill.transparency }),
      });
}

function freezeLine(line: NormalizedSimpleLine): NormalizedSimpleLine {
  return line.kind === 'none'
    ? Object.freeze({ kind: 'none' })
    : Object.freeze({
        kind: 'line',
        color: Object.freeze({ ...line.color }),
        ...(line.transparency === undefined ? {} : { transparency: line.transparency }),
        width: line.width,
        dash: line.dash,
        ...(line.cap === undefined ? {} : { cap: line.cap }),
      });
}

function validateDataLabelCompatibility(
  type: ChartType,
  options: Readonly<ChartDataLabelOptions> | undefined,
): void {
  const position = options?.position;
  if (!position) return;
  const allowed = type === 'pie' || type === 'doughnut'
    ? ['bestFit', 'center', 'insideEnd', 'outsideEnd']
    : type === 'bar' || type === 'bar3D'
      ? ['center', 'insideBase', 'insideEnd', 'outsideEnd']
      : type === 'line' || type === 'scatter' || type === 'bubble' || type === 'bubble3D'
        ? ['bottom', 'center', 'left', 'right', 'top']
        : ['center'];
  if (!allowed.includes(position)) {
    throw new TypeError(`Chart ${type} data label position ${position} is unsupported`);
  }
}

function readObject(
  value: unknown,
  allowedKeys: readonly string[],
  context: string,
): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${context} must be an object`);
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError(`${context} must be an ordinary object`);
  }
  const allowed = new Set(allowedKeys);
  const result = Object.create(null) as Record<string, unknown>;
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

function readArray(value: unknown, context: string): readonly unknown[] {
  if (!Array.isArray(value) || Object.getPrototypeOf(value) !== Array.prototype) {
    throw new TypeError(`${context} must be an ordinary array`);
  }
  const keys = Reflect.ownKeys(value);
  if (keys.length !== value.length + 1 || !keys.includes('length')) {
    throw new TypeError(`${context} must be a dense array without extra properties`);
  }
  const result: unknown[] = [];
  for (let index = 0; index < value.length; index += 1) {
    const descriptor = Object.getOwnPropertyDescriptor(value, String(index));
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`${context} item ${index} must be a data property`);
    }
    result.push(descriptor.value);
  }
  return result;
}

function containsInvalidXmlCharacter(value: string): boolean {
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    if (code === 0x9 || code === 0xa || code === 0xd) continue;
    if (code < 0x20 || code === 0xfffe || code === 0xffff) return true;
    if (code >= 0xd800 && code <= 0xdbff) {
      const next = value.charCodeAt(index + 1);
      if (next < 0xdc00 || next > 0xdfff) return true;
      index += 1;
    } else if (code >= 0xdc00 && code <= 0xdfff) {
      return true;
    }
  }
  return false;
}
