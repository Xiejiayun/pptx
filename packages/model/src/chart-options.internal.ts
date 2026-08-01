import type {
  ChartAreaOptions,
  ChartAxisOptions,
  ChartDataLabelOptions,
  ChartDataTableOptions,
  ChartFontOptions,
  ChartGroupOptions,
  ChartLegendOptions,
  ChartMarkerOptions,
  ChartOptions,
  ChartSeriesOptions,
  ChartTitleOptions,
  ChartType,
} from './chart.js';
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
  'categoryAxis',
  'valueAxis',
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
const TITLE_KEYS = [...FONT_KEYS, 'visible', 'text', 'overlay', 'rotation', 'position'] as const;
const TITLE_POSITION_KEYS = ['x', 'y'] as const;
const LEGEND_KEYS = [...FONT_KEYS, 'visible', 'position', 'overlay'] as const;
const AREA_KEYS = ['fill', 'line'] as const;
const AXIS_KEYS = [
  ...FONT_KEYS,
  'visible',
  'position',
  'title',
  'minimum',
  'maximum',
  'majorUnit',
  'minorUnit',
  'logarithmicBase',
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
const SERIES_KEYS = ['fill', 'line', 'marker'] as const;
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
  ],
  bubble: [...COMMON_GROUP_KEYS, 'scale', 'showNegativeBubbles', 'sizeRepresents'],
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
    ...optionalObject(input, 'categoryAxis', normalizeAxisOptions, 'Chart category axis'),
    ...optionalObject(input, 'valueAxis', normalizeAxisOptions, 'Chart value axis'),
    ...optionalObject(
      input,
      'secondaryCategoryAxis',
      normalizeAxisOptions,
      'Chart secondary category axis',
    ),
    ...optionalObject(
      input,
      'secondaryValueAxis',
      normalizeAxisOptions,
      'Chart secondary value axis',
    ),
    ...optionalObject(input, 'dataTable', normalizeDataTableOptions, 'Chart data table'),
    ...optionalColors(input),
    ...optionalBoolean(input, 'rightAngleAxes', 'Chart rightAngleAxes'),
    ...optionalNumber(input, 'rotationX', 'Chart rotationX', -90, 90),
    ...optionalNumber(input, 'rotationY', 'Chart rotationY', 0, 360),
    ...optionalNumber(input, 'perspective', 'Chart perspective', 0, 240),
  };
  if (
    result.categoryAxis?.minimum !== undefined
    && result.categoryAxis.maximum !== undefined
    && result.categoryAxis.minimum >= result.categoryAxis.maximum
  ) throw new RangeError('Chart category axis minimum must be less than maximum');
  if (
    result.valueAxis?.minimum !== undefined
    && result.valueAxis.maximum !== undefined
    && result.valueAxis.minimum >= result.valueAxis.maximum
  ) throw new RangeError('Chart value axis minimum must be less than maximum');
  return Object.freeze(result);
}

export function normalizeChartGroupOptions(
  type: ChartType,
  value: unknown,
  seriesCount: number,
): Readonly<ChartGroupOptions> | undefined {
  if (value === undefined) return undefined;
  const context = `Chart ${type} group options`;
  const input = readObject(value, GROUP_KEYS[type], context);
  const common: ChartGroupOptions = {
    ...optionalBoolean(input, 'varyColors', `${context} varyColors`),
    ...optionalSeriesOptions(input, type, seriesCount, context),
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
          ? optionalInteger(input, 'gapDepth', `${context} gapDepth`, 0, 500)
          : {}),
      };
      break;
    case 'bubble':
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
  });
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

function normalizeAxisOptions(value: unknown, context: string): Readonly<ChartAxisOptions> {
  const input = readObject(value, AXIS_KEYS, context);
  const result: ChartAxisOptions = {
    ...normalizeFontOptions(input, context),
    ...optionalBoolean(input, 'visible', `${context} visible`),
    ...optionalEnum(input, 'position', `${context} position`, ['bottom', 'left', 'right', 'top']),
    ...optionalObject(input, 'title', normalizeTitleOptions, `${context} title`),
    ...optionalFinite(input, 'minimum', `${context} minimum`),
    ...optionalFinite(input, 'maximum', `${context} maximum`),
    ...optionalPositive(input, 'majorUnit', `${context} majorUnit`),
    ...optionalPositive(input, 'minorUnit', `${context} minorUnit`),
    ...optionalNumber(input, 'logarithmicBase', `${context} logarithmicBase`, 2, 1_000),
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
  };
  if (
    result.minimum !== undefined
    && result.maximum !== undefined
    && result.minimum >= result.maximum
  ) throw new RangeError(`${context} minimum must be less than maximum`);
  return Object.freeze(result);
}

function normalizeDataLabelOptions(
  value: unknown,
  context: string,
): Readonly<ChartDataLabelOptions> {
  const input = readObject(value, DATA_LABEL_KEYS, context);
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

function normalizeSeriesOptions(value: unknown, context: string): Readonly<ChartSeriesOptions> {
  const input = readObject(value, SERIES_KEYS, context);
  const fill = normalizeSimpleFill(input.fill, `${context} fill`);
  const line = normalizeSimpleLine(input.line, `${context} line`);
  return Object.freeze({
    ...(fill === undefined ? {} : { fill: freezeFill(fill) }),
    ...(line === undefined ? {} : { line: freezeLine(line) }),
    ...optionalObject(input, 'marker', normalizeMarkerOptions, `${context} marker`),
  });
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
  seriesCount: number,
  context: string,
): object {
  if (!Object.hasOwn(input, 'series')) return {};
  const values = readArray(input.series, `${context} series`);
  if (values.length > seriesCount) {
    throw new RangeError(`${context} series styles cannot exceed the group series count`);
  }
  const normalized = values.map((value, index) =>
    normalizeSeriesOptions(value, `${context} series ${index}`));
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
      : type === 'line' || type === 'scatter' || type === 'bubble'
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
